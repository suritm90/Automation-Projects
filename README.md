# Automation-Projects
#Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
#cd "C:\Users\YourUserName\Documents"
#Set-Location "C:\Users\YourUserName\Documents"

#--------------------------------------------------------
# CSV FILE LOCATION
#--------------------------------------------------------

$csvPath = "C:\Users\Peram Suresh\OneDrive - TechnePlus Software India Pvt Ltd\Desktop\Pega\Resource Data.csv"

# Screenshot to be embedded at the end of every email
# Keep image004.png in the same folder as this PowerShell script.
$imagePath = Join-Path $PSScriptRoot "image004.png"

#--------------------------------------------------------
# Read CSV
#--------------------------------------------------------

$data = Import-Csv $csvPath

# Keep only Draft records
$data = $data | Where-Object {
    $_."Approval Workflow Status" -eq "Draft"
}

#--------------------------------------------------------
# Start Outlook
#--------------------------------------------------------

$outlook = New-Object -ComObject Outlook.Application

#--------------------------------------------------------
# Group by Resource
#--------------------------------------------------------

$groups = $data | Group-Object "Resource Display Name"

foreach($group in $groups)
{
    $rows = $group.Group

    $resourceName = $rows[0]."Resource Display Name"
    $resourceEmail = $rows[0]."Resource Email Address"
    #$pmEmail = $rows[0]."Project Manager Email Address"
    $pmEmails = ($rows.'Project Manager Email Address').Where({ -not [string]::IsNullOrWhiteSpace($_) }) | Select-Object -Unique
    $pmEmail = $pmEmails -join ";"


    if([string]::IsNullOrWhiteSpace($resourceEmail))
    {
        Write-Host "Skipping $resourceName"
        continue
    }

   <#$pmEmailsList = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
   foreach ($r in $rows) {
    $email = $r."Project Manager Email Address"
    if (-not [string]::IsNullOrWhiteSpace($email)) {
        [void]$pmEmailsList.Add($email.Trim())
    }
}
$pmEmail = $pmEmailsList -join ";"#>

    $firstName = $resourceName.Split(" ")[0]

    #--------------------------------------------------------
    # Build HTML Table
    #--------------------------------------------------------

    $table = "<table border='1' cellpadding='4' cellspacing='0' style='border-collapse:collapse;font-family:Calibri;font-size:11pt'>"

    $table += "<tr bgcolor='#D9EAD3'>
    <th>Resource Display Name</th>
    <th>Project Manager Email Address</th>
    <th>Resource Employee ID</th>
    <th>Resource Email Address</th>
    <th>Project Manager</th>
    <th>Direct Manager</th>
    <th>Invoice #</th>
    <th>TimeSheetID</th>
    <th>Approval Workflow Status</th>
    <th>Work Date</th>
    <th>Hours</th>
    <th>Missing Hours</th>
    <th>Project Code</th>
    <th>Project Name</th>
    <th>Project Role Name</th>
    <th>Billing Adjusted Hourly Rate</th>
    <th>Billing Adjusted Daily Rate</th>
    <th>Billing Adjusted Revenue</th>
    <th>Engagement Currency Code</th>
    <th>Resource ResourceClassificationCode</th>
    </tr>"

    foreach($r in $rows)
    {
        $table += "<tr>"
        $table += "<td>$($r.'Resource Display Name')</td>"
        $table += "<td>$($r.'Project Manager Email Address')</td>"
        $table += "<td>$($r.'Resource Employee ID')</td>"
        $table += "<td>$($r.'Resource Email Address')</td>"
        $table += "<td>$($r.'Project Manager')</td>"
        $table += "<td>$($r.'Direct Manager')</td>"
        $table += "<td>$($r.'Invoice #')</td>"
        $table += "<td>$($r.'TimeSheetID')</td>"
        $table += "<td>$($r.'Approval Workflow Status')</td>"
        $table += "<td>$($r.'Work Date')</td>"
        $table += "<td>$($r.Hours)</td>"
        $table += "<td>$($r.'Missing Hours')</td>"
        $table += "<td>$($r.'Project Code')</td>"
        $table += "<td>$($r.'Project Name')</td>"
        $table += "<td>$($r.'Project Role Name')</td>"
        $table += "<td>$($r.'Billing Adjusted Hourly Rate')</td>"
        $table += "<td>$($r.'Billing Adjusted Daily Rate')</td>"
        $table += "<td>$($r.'Billing Adjusted Revenue')</td>"
        $table += "<td>$($r.'Engagement Currency Code')</td>"
        $table += "<td>$($r.'Resource ResourceClassificationCode')</td>"
        $table += "</tr>"
    }

    $table += "</table>"

    #--------------------------------------------------------
    # Email Body
    #--------------------------------------------------------

    $body = @"
<html>

<body style='font-family:Calibri;font-size:11pt'>

<p>Hello $firstName,</p>

<p>
We see that for the below dates you have entered hours against a project but never submitted till date for approval so these are in Draft stage.
</p>

<p>
If these are legitimate hours then please submit these for approval.
If not, kindly delete the draft timecards from Projector.
</p>

<p>
This is very critical from a cash flow standpoint and requires your immediate attention.
</p>

$table

<br>

<b>Steps to delete Draft Timecards</b>

<ol>

<li>Open Projector Web</li>

<li>Click Enter Time</li>

<li>Select your name and week</li>

<li>Delete the draft timecards OR submit them</li>

<li>Click Save</li>

</ol>

<br>

Regards,<br>

PMO Team<br><br>

<img src="cid:ProjectorScreenshot" alt="Projector Timesheet Screenshot" style="max-width:100%;">

</body>

</html>
"@

    #--------------------------------------------------------
    # Create Email
    #--------------------------------------------------------

    $mail = $outlook.CreateItem(0)

    $mail.To = $resourceEmail
    $mail.CC = $pmEmail
    $mail.Subject = "Action Required - Draft Timecards Pending Approval"
    $mail.HTMLBody = $body

    #--------------------------------------------------------
    # Embed Projector screenshot at the end of the email
    #--------------------------------------------------------

    if(Test-Path $imagePath)
    {
        # 1 = Outlook attachment type: olByValue
        $attachment = $mail.Attachments.Add($imagePath, 1, 0, "Projector Timesheet Screenshot")

        # Set Content-ID so the image is displayed inline in the HTML body
        $contentIdProperty = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
        $attachment.PropertyAccessor.SetProperty($contentIdProperty, "ProjectorScreenshot")
    }
    else
    {
        Write-Warning "Screenshot not found: $imagePath"
    }

    # Review before sending
    #$mail.Display()

    # To send automatically instead, comment the line above
    # and uncomment the line below.
    $mail.Send()

    Write-Host "Prepared email for $resourceName"
}

Write-Host "Completed Sending Follow-up emails to all the resources."
