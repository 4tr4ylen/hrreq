param(
  [string]$SiteUrl = "https://atranox.sharepoint.com/sites/test",
  [string]$ListTitle = "HR Requests",
  [string[]]$Departments = @("HR","IT","Finance","Sales","Operations"),
  [string[]]$RequestTypes = @("Leave Request","Equipment Request","Policy Question","Benefits Question","Other")
)

# Requires: PnP.PowerShell (Install-Module PnP.PowerShell -Scope CurrentUser)
Import-Module PnP.PowerShell -ErrorAction Stop

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive

# Ensure list exists
$list = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
if (-not $list) {
  Write-Host "Creating list '$ListTitle'" -ForegroundColor Cyan
  $list = New-PnPList -Title $ListTitle -Template GenericList -OnQuickLaunch
}

# Enable attachments and versioning
Write-Host "Configuring attachments and versioning" -ForegroundColor Cyan
Set-PnPList -Identity $ListTitle -EnableAttachments:$true -EnableVersioning:$true -MajorVersions 500 | Out-Null

# Helper: ensure field exists
function Ensure-Field {
  param(
    [string]$InternalName,
    [string]$DisplayName,
    [string]$Type,
    [bool]$Required = $false,
    [bool]$AddToDefaultView = $true
  )
  $field = Get-PnPField -List $ListTitle -Identity $InternalName -ErrorAction SilentlyContinue
  if (-not $field) {
    Write-Host "Adding field $DisplayName ($InternalName) of type $Type" -ForegroundColor Green
    Add-PnPField -List $ListTitle -DisplayName $DisplayName -InternalName $InternalName -Type $Type -AddToDefaultView:$AddToDefaultView -Required:$Required | Out-Null
  } else {
    if ($field.Required -ne $Required) {
      Set-PnPField -List $ListTitle -Identity $InternalName -Required:$Required | Out-Null
    }
  }
}

# Helper: add choice field if missing
function Ensure-ChoiceField {
  param(
    [string]$InternalName,
    [string]$DisplayName,
    [string[]]$Choices,
    [string]$DefaultValue = $null,
    [bool]$Required = $false
  )
  $field = Get-PnPField -List $ListTitle -Identity $InternalName -ErrorAction SilentlyContinue
  if (-not $field) {
    Write-Host "Adding choice field $DisplayName ($InternalName)" -ForegroundColor Green
    Add-PnPField -List $ListTitle -DisplayName $DisplayName -InternalName $InternalName -Type Choice -AddToDefaultView -Required:$Required -Choices $Choices -DefaultValue $DefaultValue | Out-Null
  }
}

# Columns
Ensure-Field -InternalName "Title" -DisplayName "Title" -Type "Text" -Required $true | Out-Null
Ensure-ChoiceField -InternalName "RequestType" -DisplayName "Request Type" -Choices $RequestTypes -DefaultValue "Other" -Required $true

# Description (Note)
$desc = Get-PnPField -List $ListTitle -Identity "Description" -ErrorAction SilentlyContinue
if (-not $desc) {
  Write-Host "Adding Description (Note)" -ForegroundColor Green
  Add-PnPField -List $ListTitle -DisplayName "Description" -InternalName "Description" -Type Note -AddToDefaultView -Required:$true | Out-Null
}

# Department (Choice) – configurable values
if ($Departments -and $Departments.Count -gt 0) {
  Ensure-ChoiceField -InternalName "Department" -DisplayName "Department" -Choices $Departments -Required $true
} else {
  Ensure-Field -InternalName "Department" -DisplayName "Department" -Type Text -Required $true | Out-Null
}

# Person fields
Ensure-Field -InternalName "Requestor" -DisplayName "Requestor" -Type User -Required $false | Out-Null
Ensure-Field -InternalName "Manager" -DisplayName "Manager" -Type User -Required $false | Out-Null

# Status & approval
Ensure-ChoiceField -InternalName "Status" -DisplayName "Status" -Choices @("Draft","Submitted","Pending Approval","Approved","Rejected","Completed") -DefaultValue "Submitted" -Required $true
Ensure-ChoiceField -InternalName "ApprovalOutcome" -DisplayName "Approval Outcome" -Choices @("Approved","Rejected") -Required $false

# Approver Comments (Note)
$ac = Get-PnPField -List $ListTitle -Identity "ApproverComments" -ErrorAction SilentlyContinue
if (-not $ac) {
  Write-Host "Adding Approver Comments (Note)" -ForegroundColor Green
  Add-PnPField -List $ListTitle -DisplayName "Approver Comments" -InternalName "ApproverComments" -Type Note -AddToDefaultView | Out-Null
}

# Add useful fields to default view (idempotent)
Write-Host "Ensuring default view shows key fields" -ForegroundColor Cyan
$view = Get-PnPView -List $ListTitle -Identity "All Items"
$desired = @("Title","RequestType","Department","Status","Requestor","Manager","Modified")
foreach ($f in $desired) {
  try { Add-PnPViewField -List $ListTitle -Identity $view -Field $f -ErrorAction SilentlyContinue | Out-Null } catch {}
}

Write-Host "List '$ListTitle' is configured." -ForegroundColor Green

