<#
Author: Harish C M
Provide feedback at harishcm21@gmail.com
#>

if($FormatEnumerationLimit -ne "-1")
{
$FormatEnumerationLimit = -1
}
if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript") {
    $presentpath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
}
else {
    $presentpath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0])
}
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
#icon
$iconimageBytes = [Convert]::FromBase64String("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")
$ims = New-Object IO.MemoryStream($iconimageBytes, 0, $iconimageBytes.Length)
$ims.Write($iconimageBytes, 0, $iconimageBytes.Length);

#URL used
$accessURL="https://login.microsoftonline.com"
$GraphURL="https://graph.microsoft.com/v1.0"


#M365 main form

$Form = new-object system.windows.forms.form -property @{
    Icon            = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ims).GetHIcon())
    MaximizeBox     = $false
    MinimizeBox     = $false
    Text            = "Microsoft 365"
   # AutoSize        = $True
    Size            = New-Object System.Drawing.Size(450, 430)
    StartPosition   = "CenterScreen"
    FormBorderStyle = 'Fixed3D'
    Topmost         = $false
    CancelButton    = $Cancel
}

#group box for authentication
$AuthGP = new-object system.windows.forms.GroupBox -Property @{
    Visible  = $true
    Enabled  = $true
    Location = new-object system.Drawing.point(10, 10)
    Size     = new-object system.drawing.size(420, 155)
    Text     = 'Authentication'
}

#application/client id text box

$global:APPId = New-Object System.Windows.Forms.TextBox -Property @{
    Location        = new-object system.Drawing.point(180, 30)
    TabIndex        = 9
    Size            = new-object system.drawing.size(230, 20)
    add_TextChanged = {
           $Validate.Enabled = ![string]::IsNullOrWhiteSpace($global:APPId.Text) -and ![string]::IsNullOrWhiteSpace($global:ClientSecret.Text) -and ![string]::IsNullOrWhiteSpace($global:TenantId.Text) }
}

#client secret text box

$global:ClientSecret = New-Object System.Windows.Forms.TextBox -Property @{
    Location        = new-object system.Drawing.point(180, 60)
    Size            = new-object system.drawing.size(230, 20)
    add_TextChanged = {
       $Validate.Enabled = ![string]::IsNullOrWhiteSpace($global:APPId.Text) -and ![string]::IsNullOrWhiteSpace($global:ClientSecret.Text) -and ![string]::IsNullOrWhiteSpace($global:TenantId.Text) }
}

#tenant Id text box

$global:TenantId = New-Object System.Windows.Forms.TextBox -Property @{
    Location        = new-object system.Drawing.point(180, 90)
    #TabIndex        = 11
    Size            = new-object system.drawing.size(230, 20)
    add_TextChanged = {
        $Validate.Enabled = ![string]::IsNullOrWhiteSpace($global:APPId.Text) -and ![string]::IsNullOrWhiteSpace($global:ClientSecret.Text) -and ![string]::IsNullOrWhiteSpace($global:TenantId.Text) }
}

#validation of api credentials
$Validate = New-Object System.Windows.Forms.Button -Property @{
    Location        = new-object System.Drawing.Point(220, 115)
    Enabled         = $false
    Text            ="Validate"
    Size            = new-object system.drawing.size(150, 30)
    add_Click       = {
    try
    {
    $header= @{"Content-Type"="application/x-www-form-urlencoded"}
    $Body = @{grant_type = "client_credentials"; scope = "https://graph.microsoft.com/.default"; client_id = $($APPId.text); client_secret = $($ClientSecret.text) }
      $global:token=Invoke-RestMethod -Method Post -Uri $accessURL/$($TenantId.Text)/oauth2/v2.0/token -Body $Body -Headers $header #https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token
      $global:Graphheader = @{Authorization = "$($token.token_type) $($token.access_token)" }
      $ReportsGP.Enabled=$True
      [System.Windows.Forms.MessageBox]::Show("Validation successful.", "Information", "OK", "info") | out-null
              }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Following error occurred while authenticating.`n $_", "Error", "OK", "Error") | out-null
            return
        }
    }
    }

#add values to AuthGP
$AuthGP.Controls.AddRange(@(
                        $global:APPId,
                        $global:ClientSecret,
                        $global:TenantId,
                        $Validate
                        new-object System.Windows.Forms.Label -Property @{
                            Text      = 'Application/Client Id:'
                            Location  = New-Object System.Drawing.Point(20, 30)
                            Size     = new-object system.drawing.size(150, 20)
                            textalign = [System.Drawing.ContentAlignment]::MiddleRight
                        }
                        new-object System.Windows.Forms.Label -Property @{
                            Text      = 'Client Secret:'
                            Location  = New-Object System.Drawing.Point(20, 60)
                            Size     = new-object system.drawing.size(150, 20)
                            textalign = [System.Drawing.ContentAlignment]::MiddleRight
                        }
                        new-object System.Windows.Forms.Label -Property @{
                            Text      = 'Tenant Id:'
                            Location  = New-Object System.Drawing.Point(20, 90)
                            Size     = new-object system.drawing.size(150, 20)
                            textalign = [System.Drawing.ContentAlignment]::MiddleRight
                        }
                    ))

#group box for reports
$ReportsGP = new-object system.windows.forms.GroupBox -Property @{
    Visible  = $true
    Enabled  = $false
    Location = new-object system.Drawing.point(10, 180)
    Size     = new-object system.drawing.size(420, 200)
    Text     = 'Graph API Reports'
}

#Combobox for Graph API type
$Type = new-object System.Windows.Forms.ComboBox -Property @{
    Location                 = new-object System.Drawing.Point(180, 30)
    Size     = new-object system.drawing.size(230, 20)
    DropDownStyle            = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    Enabled                   = $true
    add_SelectedIndexChanged = {
        switch ($Type.SelectedItem) {
            "Reports" {$SubTypevalue=@("Azure AD activity"); $SubType.Enabled=$true;$ViewReportGenerate.Enabled=$false }
            "Security" {$SubTypevalue=@("Alerts and incidents"); $SubType.Enabled=$true;$ViewReportGenerate.Enabled=$false }
             Default { }
        }
        $SubType.Items.Clear()
        $ReportsType.Items.Clear()
        $PermissionLabel.Text="Required permission"
        $SubType.Items.Addrange(@(
        $SubTypevalue
    ))
    }
}
#adding subsription item to combobox
$Type.Items.Addrange(@(
        "Reports",
        "Security"
    ))

#Combobox for Graph API subtype
$SubType = new-object System.Windows.Forms.ComboBox -Property @{
    Location                 = new-object System.Drawing.Point(180, 60)
    Size     = new-object system.drawing.size(230, 20)
    DropDownStyle            = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    Enabled                   = $false
    add_SelectedIndexChanged = {
        switch ($SubType.SelectedItem) {
            "Azure AD activity" {$ReportsTypeValue=@("Sign in","Directory Audit","User Registration Details");$ReportsType.Enabled=$true;$ViewReportGenerate.Enabled=$false }
            "Alerts and incidents" {$ReportsTypeValue=@("Alerts","Incidents");$ReportsType.Enabled=$true;$ViewReportGenerate.Enabled=$false }
             Default { }
        }
                 $ReportsType.Items.Clear()
                 $PermissionLabel.Text="Required permission"
         $ReportsType.Items.Addrange(@(
        $ReportsTypeValue
    ))
    }
}

#Combobox for Graph API entity
$ReportsType = new-object System.Windows.Forms.ComboBox -Property @{
    Location                 = new-object System.Drawing.Point(180, 90)
    Size     = new-object system.drawing.size(230, 20)
    DropDownStyle            = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    Enabled                   = $false
    add_SelectedIndexChanged = {
        switch ($ReportsType.SelectedItem) {
            "Sign in" {$ViewReportGenerate.Enabled=$True; $global:Callmethod="GET"; $global:CallURL="/auditLogs/signIns"; $Global:activity=$null;$PermissionLabel.Text="Required permission is AuditLog.Read.All and Directory.Read.All" }
            "Directory Audit" {$ViewReportGenerate.Enabled=$True; $global:Callmethod="GET"; $global:CallURL="/auditLogs/directoryAudits"; $Global:activity=$null; $PermissionLabel.Text="Required permission is AuditLog.Read.All and Directory.Read.All"  }
            "User Registration Details" {$ViewReportGenerate.Enabled=$True; $global:Callmethod="GET"; $global:CallURL="/reports/authenticationMethods/userRegistrationDetails"; $Global:activity=$null; $PermissionLabel.Text="Required permission is AuditLog.Read.All"}
            "Alerts" {$ViewReportGenerate.Enabled=$True; $global:Callmethod="GET"; $global:CallURL="/security/alerts_v2"; $Global:activity=$null; $PermissionLabel.Text="Required permission is SecurityAlert.Read.All"  }
            "Incidents" {$ViewReportGenerate.Enabled=$True; $global:Callmethod="GET"; $global:CallURL="/security/incidents"; $Global:activity=$null; $PermissionLabel.Text="Required permission is SecurityIncident.Read.All"  }
             Default { }
        }
    }
}

#View Report generator
$ViewReportGenerate = New-Object System.Windows.Forms.Button -Property @{
    Location        = new-object System.Drawing.Point(220, 160)
    Enabled         = $false
    Text            ="View Report"
    Size            = new-object system.drawing.size(150, 30)
    add_Click       = {
    try{
    $ViewReportGenerate.Enabled=$false
    $Global:activity=Invoke-RestMethod -Uri "$($GraphURL)$($CallURL)" -Method $Callmethod -Headers $Graphheader
    ($Global:activity.value | ConvertTo-Json | ConvertFrom-Json) | Out-GridView -Title "$($GraphURL)$($CallURL)"
    $ViewReportGenerate.Enabled=$true
    }
    catch
    {
    [System.Windows.Forms.MessageBox]::Show($_,"M365","OK","Error")
    }

    }
    }
$PermissionLabel=  new-object System.Windows.Forms.Label -Property @{
                            Text      = 'Required permission'
                            Location  = New-Object System.Drawing.Point(10, 120)
                            Size     = new-object system.drawing.size(400, 30)
                            textalign = [System.Drawing.ContentAlignment]::MiddleCenter
                        }

#add values to AuthGP
$ReportsGP.Controls.AddRange(@(
                        $Type,
                        $SubType,
                        $ReportsType,
                      #  $ReportGenerate,
                        $ViewReportGenerate,
                        $PermissionLabel
                        new-object System.Windows.Forms.Label -Property @{
                            Text      = 'Graph API Types:'
                            Location  = New-Object System.Drawing.Point(20, 30)
                            Size     = new-object system.drawing.size(150, 20)
                            textalign = [System.Drawing.ContentAlignment]::MiddleRight
                        }
                        new-object System.Windows.Forms.Label -Property @{
                            Text      = 'Graph API Subtypes:'
                            Location  = New-Object System.Drawing.Point(20, 60)
                            Size     = new-object system.drawing.size(150, 20)
                            textalign = [System.Drawing.ContentAlignment]::MiddleRight
                        }
                        new-object System.Windows.Forms.Label -Property @{
                            Text      = 'Graph API Reports:'
                            Location  = New-Object System.Drawing.Point(20, 90)
                            Size     = new-object system.drawing.size(150, 20)
                            textalign = [System.Drawing.ContentAlignment]::MiddleRight
                        }
                    ))

#add controls to form
$Form.Controls.AddRange(@(
        $AuthGP,
        $ReportsGP
    ))

#display the form
$Form.ShowDialog() | out-null