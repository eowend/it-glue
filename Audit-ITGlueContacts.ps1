<#
.SYNOPSIS
Audits IT Glue contact records for common data formatting issues.

.DESCRIPTION
This PowerShell script performs a read-only audit of contact records stored
in IT Glue.

It retrieves contacts through the IT Glue REST API and generates a CSV
report containing records that may require manual normalization.

The script does not update, delete, or create objects in IT Glue. All
findings must be reviewed and corrected manually through the IT Glue web
interface or another approved process.

A graphical interface prompts the operator for:

- IT Glue API key
- IT Glue API region
- CSV report output location
- Whether the report should open after completion

.AUDIT RULES
The script checks the following contact properties:

1. Email capitalization

   Each email address is compared against its lowercase equivalent.

   Example detected:

       JOrtiz@example.com

   Suggested value:

       jortiz@example.com

2. Non-ASCII contact field characters

   The following fields are reviewed:

   - First Name
   - Last Name
   - Title

   Printable ASCII characters from U+0020 through U+007E are accepted.

   This includes:

   - English letters
   - Numbers
   - Standard spaces
   - Standard punctuation
   - Common keyboard symbols

   Characters outside that range are reported because they may cause
   compatibility issues with external systems, exports, synchronization
   tools, APIs, or other integrations.

   Examples detected:

       Jose
       Munoz
       Kommerling
       O'Neill using a curved apostrophe
       Titles using long dashes or special spaces

   The report includes:

   - The affected field
   - The current value
   - The detected Unicode character
   - The Unicode code point
   - A suggested printable ASCII value when one can be generated

3. Telephone number prefix

   Every documented telephone number must begin with a plus sign.

   Valid examples:

       +1 (787) 371-8815
       +57 300 123 4567
       +44-20-1234-5678
       +17873718815

   Examples reported:

       787-371-8815
       1 (787) 371-8815
       0017873718815

   The script does not enforce a specific spacing, country code length,
   area code format, or separator format. It only verifies that the
   documented value begins with the plus sign.

.OUTPUT
The script creates a CSV report containing one row per detected issue.

Depending on the issue type, the report may include:

- Organization
- Contact name
- Contact email
- Issue type
- Affected field
- Current value
- Suggested value
- Invalid characters
- Email label
- Phone label
- Phone extension
- Primary field status
- Direct IT Glue contact URL
- IT Glue contact ID

A single contact may appear more than once when multiple issues are found.

.API BEHAVIOR
The script uses GET requests against the IT Glue contacts endpoint.

Contact retrieval is paginated to ensure that all available contacts are
reviewed.

Temporary API failures and rate-limit responses may be retried before the
audit stops with an error.

.SECURITY
The API key is entered through the graphical interface.

The script does not intentionally:

- Save the API key to disk
- Write the API key to the CSV report
- Write the API key to the PowerShell console
- Store the API key in an environment variable
- Modify IT Glue contact records

The API key should have only the permissions required to read contact
records.

.REQUIREMENTS
- Windows operating system
- Windows PowerShell 5.1 or PowerShell 7 on Windows
- System.Windows.Forms
- System.Drawing
- Network access to the selected IT Glue API region
- Valid IT Glue API key with contact read permissions

.SUPPORTED REGIONS
- US
- EU
- AU

.USAGE
Run the script directly from PowerShell:

    & '.\Audit-ITGlueContacts.ps1'

Do not dot-source the script unless its functions and variables must be
imported into the current PowerShell session.

Recommended execution:

    & 'C:\Path\Audit-ITGlueContacts.ps1'

.SAFETY
This is an audit-only script.

It does not send PATCH, POST, PUT, or DELETE requests to IT Glue.

Review all suggested values before making manual changes. Suggested values
are provided for convenience and may not represent the preferred spelling
or formatting for every organization.

.AUTHOR
Juan Ortiz
juan.ortiz@toconnect.net

.VERSION
3.1

.CHANGELOG
Version 3.1
- Added telephone number validation.
- Telephone numbers are reported when they do not begin with a plus sign.
- Maintained email capitalization auditing.
- Maintained non-ASCII character auditing.
- Maintained graphical configuration and CSV reporting.
- Maintained read-only IT Glue API behavior.

Version 3.0
- Added telephone number format auditing.
- Added suggested telephone number formatting.

Version 2.0
- Added non-ASCII character detection for contact fields.
- Added printable ASCII normalization suggestions.

Version 1.0
- Added email capitalization auditing.
- Added graphical API and report configuration.

#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# Any character outside printable ASCII is reported.
$script:InvalidCharacterPattern = '[^\x20-\x7E]'

# A documented phone number is valid when its displayed value begins with +.
$script:PhoneNumberPattern = '^\+'

function Show-ConfigurationDialog {
    [CmdletBinding()]
    param()

    $Timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $DefaultFileName = "ITGlue-ContactDataAudit-$Timestamp.csv"
    $DefaultFolder = [Environment]::GetFolderPath(
        [Environment+SpecialFolder]::Desktop
    )
    $DefaultReportPath = Join-Path -Path $DefaultFolder -ChildPath $DefaultFileName

    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = 'IT Glue Contact Data Audit'
    $Form.Size = New-Object System.Drawing.Size(680, 365)
    $Form.StartPosition = 'CenterScreen'
    $Form.FormBorderStyle = 'FixedDialog'
    $Form.MaximizeBox = $false
    $Form.MinimizeBox = $false
    $Form.TopMost = $true

    $TitleLabel = New-Object System.Windows.Forms.Label
    $TitleLabel.Text = 'IT Glue Contact Data Audit'
    $TitleLabel.Font = New-Object System.Drawing.Font(
        'Segoe UI',
        14,
        [System.Drawing.FontStyle]::Bold
    )
    $TitleLabel.AutoSize = $true
    $TitleLabel.Location = New-Object System.Drawing.Point(20, 15)
    [void]$Form.Controls.Add($TitleLabel)

    $DescriptionLabel = New-Object System.Windows.Forms.Label
    $DescriptionLabel.Text = (
        'Find uppercase email addresses, non-ASCII contact data, and phone ' +
        'numbers that do not begin with +. No IT Glue objects will be modified.'
    )
    $DescriptionLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $DescriptionLabel.AutoSize = $true
    $DescriptionLabel.Location = New-Object System.Drawing.Point(22, 50)
    [void]$Form.Controls.Add($DescriptionLabel)

    $ApiKeyLabel = New-Object System.Windows.Forms.Label
    $ApiKeyLabel.Text = 'IT Glue API key'
    $ApiKeyLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $ApiKeyLabel.AutoSize = $true
    $ApiKeyLabel.Location = New-Object System.Drawing.Point(22, 88)
    [void]$Form.Controls.Add($ApiKeyLabel)

    $ApiKeyTextBox = New-Object System.Windows.Forms.TextBox
    $ApiKeyTextBox.Location = New-Object System.Drawing.Point(160, 84)
    $ApiKeyTextBox.Size = New-Object System.Drawing.Size(480, 25)
    $ApiKeyTextBox.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $ApiKeyTextBox.UseSystemPasswordChar = $true
    [void]$Form.Controls.Add($ApiKeyTextBox)

    $RegionLabel = New-Object System.Windows.Forms.Label
    $RegionLabel.Text = 'IT Glue region'
    $RegionLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $RegionLabel.AutoSize = $true
    $RegionLabel.Location = New-Object System.Drawing.Point(22, 128)
    [void]$Form.Controls.Add($RegionLabel)

    $RegionComboBox = New-Object System.Windows.Forms.ComboBox
    $RegionComboBox.Location = New-Object System.Drawing.Point(160, 124)
    $RegionComboBox.Size = New-Object System.Drawing.Size(180, 25)
    $RegionComboBox.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $RegionComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$RegionComboBox.Items.Add('US')
    [void]$RegionComboBox.Items.Add('EU')
    [void]$RegionComboBox.Items.Add('AU')
    $RegionComboBox.SelectedIndex = 0
    [void]$Form.Controls.Add($RegionComboBox)

    $ReportLabel = New-Object System.Windows.Forms.Label
    $ReportLabel.Text = 'CSV report'
    $ReportLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $ReportLabel.AutoSize = $true
    $ReportLabel.Location = New-Object System.Drawing.Point(22, 168)
    [void]$Form.Controls.Add($ReportLabel)

    $ReportPathTextBox = New-Object System.Windows.Forms.TextBox
    $ReportPathTextBox.Location = New-Object System.Drawing.Point(160, 164)
    $ReportPathTextBox.Size = New-Object System.Drawing.Size(395, 25)
    $ReportPathTextBox.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $ReportPathTextBox.Text = $DefaultReportPath
    [void]$Form.Controls.Add($ReportPathTextBox)

    $BrowseButton = New-Object System.Windows.Forms.Button
    $BrowseButton.Text = 'Browse...'
    $BrowseButton.Location = New-Object System.Drawing.Point(565, 162)
    $BrowseButton.Size = New-Object System.Drawing.Size(75, 28)
    $BrowseButton.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    [void]$Form.Controls.Add($BrowseButton)

    $OpenReportCheckBox = New-Object System.Windows.Forms.CheckBox
    $OpenReportCheckBox.Text = 'Open the CSV report when completed'
    $OpenReportCheckBox.Location = New-Object System.Drawing.Point(160, 205)
    $OpenReportCheckBox.Size = New-Object System.Drawing.Size(320, 25)
    $OpenReportCheckBox.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $OpenReportCheckBox.Checked = $true
    [void]$Form.Controls.Add($OpenReportCheckBox)

    $RulesLabel = New-Object System.Windows.Forms.Label
    $RulesLabel.Text = (
        'Allowed text: printable ASCII. Phone numbers must begin with +.'
    )
    $RulesLabel.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $RulesLabel.AutoSize = $true
    $RulesLabel.Location = New-Object System.Drawing.Point(160, 235)
    [void]$Form.Controls.Add($RulesLabel)

    $RunButton = New-Object System.Windows.Forms.Button
    $RunButton.Text = 'Run Audit'
    $RunButton.Location = New-Object System.Drawing.Point(460, 275)
    $RunButton.Size = New-Object System.Drawing.Size(95, 32)
    $RunButton.Font = New-Object System.Drawing.Font(
        'Segoe UI',
        9,
        [System.Drawing.FontStyle]::Bold
    )
    [void]$Form.Controls.Add($RunButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Text = 'Cancel'
    $CancelButton.Location = New-Object System.Drawing.Point(565, 275)
    $CancelButton.Size = New-Object System.Drawing.Size(75, 32)
    $CancelButton.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    [void]$Form.Controls.Add($CancelButton)

    $Form.AcceptButton = $RunButton
    $Form.CancelButton = $CancelButton

    $BrowseButton.Add_Click({
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog

        try {
            $SaveFileDialog.Title = 'Select the CSV report location'
            $SaveFileDialog.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
            $SaveFileDialog.DefaultExt = 'csv'
            $SaveFileDialog.AddExtension = $true
            $SaveFileDialog.OverwritePrompt = $true

            $CurrentPath = $ReportPathTextBox.Text.Trim()

            if (-not [string]::IsNullOrWhiteSpace($CurrentPath)) {
                try {
                    $CurrentFolder = Split-Path -Path $CurrentPath -Parent
                    $CurrentFile = Split-Path -Path $CurrentPath -Leaf

                    if (
                        -not [string]::IsNullOrWhiteSpace($CurrentFolder) -and
                        (Test-Path -LiteralPath $CurrentFolder -PathType Container)
                    ) {
                        $SaveFileDialog.InitialDirectory = $CurrentFolder
                    }

                    if (-not [string]::IsNullOrWhiteSpace($CurrentFile)) {
                        $SaveFileDialog.FileName = $CurrentFile
                    }
                }
                catch {
                    $SaveFileDialog.InitialDirectory = $DefaultFolder
                    $SaveFileDialog.FileName = $DefaultFileName
                }
            }

            if (
                $SaveFileDialog.ShowDialog() -eq
                [System.Windows.Forms.DialogResult]::OK
            ) {
                $ReportPathTextBox.Text = $SaveFileDialog.FileName
            }
        }
        finally {
            $SaveFileDialog.Dispose()
        }
    })

    $RunButton.Add_Click({
        $ApiKey = $ApiKeyTextBox.Text.Trim()
        $Region = [string]$RegionComboBox.SelectedItem
        $ReportPath = $ReportPathTextBox.Text.Trim()

        if ([string]::IsNullOrWhiteSpace($ApiKey)) {
            [System.Windows.Forms.MessageBox]::Show(
                'Enter the IT Glue API key.',
                'Missing API Key',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null

            $ApiKeyTextBox.Focus()
            return
        }

        if ([string]::IsNullOrWhiteSpace($Region)) {
            [System.Windows.Forms.MessageBox]::Show(
                'Select the IT Glue region.',
                'Missing Region',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            return
        }

        if ([string]::IsNullOrWhiteSpace($ReportPath)) {
            [System.Windows.Forms.MessageBox]::Show(
                'Select a location for the CSV report.',
                'Missing Report Path',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null

            $ReportPathTextBox.Focus()
            return
        }

        if ([System.IO.Path]::GetExtension($ReportPath) -ne '.csv') {
            $ReportPath = "$ReportPath.csv"
            $ReportPathTextBox.Text = $ReportPath
        }

        try {
            $ReportFolder = Split-Path -Path $ReportPath -Parent
        }
        catch {
            $ReportFolder = $null
        }

        if (
            [string]::IsNullOrWhiteSpace($ReportFolder) -or
            -not (Test-Path -LiteralPath $ReportFolder -PathType Container)
        ) {
            [System.Windows.Forms.MessageBox]::Show(
                'The selected report folder does not exist.',
                'Invalid Report Path',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null

            $ReportPathTextBox.Focus()
            return
        }

        if (Test-Path -LiteralPath $ReportPath -PathType Leaf) {
            $OverwriteResponse = [System.Windows.Forms.MessageBox]::Show(
                "The report already exists.`r`n`r`nDo you want to overwrite it?",
                'Confirm Overwrite',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )

            if ($OverwriteResponse -ne [System.Windows.Forms.DialogResult]::Yes) {
                return
            }
        }

        $Form.Tag = [PSCustomObject]@{
            ApiKey     = $ApiKey
            Region     = $Region
            ReportPath = $ReportPath
            OpenReport = [bool]$OpenReportCheckBox.Checked
        }

        $Form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $Form.Close()
    })

    $ApiKeyTextBox.Focus()
    $DialogResult = $Form.ShowDialog()

    if (
        $DialogResult -ne [System.Windows.Forms.DialogResult]::OK -or
        $null -eq $Form.Tag
    ) {
        $Form.Dispose()
        return $null
    }

    $Result = $Form.Tag
    $Form.Dispose()
    return $Result
}

function Invoke-ITGlueGetRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Uri,

        [Parameter(Mandatory)]
        [hashtable]$Headers,

        [ValidateRange(1, 10)]
        [int]$MaximumAttempts = 5
    )

    for ($Attempt = 1; $Attempt -le $MaximumAttempts; $Attempt++) {
        try {
            return Invoke-RestMethod `
                -Method Get `
                -Uri $Uri `
                -Headers $Headers `
                -ErrorAction Stop
        }
        catch {
            $StatusCode = $null

            try {
                $StatusCode = [int]$_.Exception.Response.StatusCode
            }
            catch {
                $StatusCode = $null
            }

            $RetryableStatusCodes = @(429, 500, 502, 503, 504)

            if (
                $StatusCode -in $RetryableStatusCodes -and
                $Attempt -lt $MaximumAttempts
            ) {
                $DelaySeconds = $Attempt * 5
                Write-Warning "IT Glue API returned HTTP $StatusCode. Retrying in $DelaySeconds seconds. Attempt $Attempt of $MaximumAttempts."
                Start-Sleep -Seconds $DelaySeconds
                continue
            }

            throw
        }
    }
}

function Test-ContainsInvalidCharacters {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    if ([string]::IsNullOrEmpty($Value)) {
        return $false
    }

    return [regex]::IsMatch($Value, $script:InvalidCharacterPattern)
}

function Get-InvalidCharacterList {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    if ([string]::IsNullOrEmpty($Value)) {
        return ''
    }

    $Matches = [regex]::Matches($Value, $script:InvalidCharacterPattern)

    if ($Matches.Count -eq 0) {
        return ''
    }

    $Results = [System.Collections.Generic.List[string]]::new()
    $ProcessedCodes = [System.Collections.Generic.HashSet[int]]::new()

    foreach ($Match in $Matches) {
        $Character = [string]$Match.Value
        $CharacterCode = [int][char]$Character[0]

        if (-not $ProcessedCodes.Add($CharacterCode)) {
            continue
        }

        $UnicodeCode = 'U+' + $CharacterCode.ToString('X4')

        $CharacterName = switch ($CharacterCode) {
            0x0009 { 'TAB' }
            0x000A { 'LINE FEED' }
            0x000D { 'CARRIAGE RETURN' }
            0x00A0 { 'NON-BREAKING SPACE' }
            0x00AD { 'SOFT HYPHEN' }
            0x200B { 'ZERO-WIDTH SPACE' }
            0x2010 { 'HYPHEN' }
            0x2011 { 'NON-BREAKING HYPHEN' }
            0x2012 { 'FIGURE DASH' }
            0x2013 { 'EN DASH' }
            0x2014 { 'EM DASH' }
            0x2015 { 'HORIZONTAL BAR' }
            0x2018 { 'LEFT SINGLE QUOTATION MARK' }
            0x2019 { 'RIGHT SINGLE QUOTATION MARK' }
            0x201A { 'SINGLE LOW-9 QUOTATION MARK' }
            0x201B { 'SINGLE HIGH-REVERSED-9 QUOTATION MARK' }
            0x201C { 'LEFT DOUBLE QUOTATION MARK' }
            0x201D { 'RIGHT DOUBLE QUOTATION MARK' }
            0x201E { 'DOUBLE LOW-9 QUOTATION MARK' }
            0x201F { 'DOUBLE HIGH-REVERSED-9 QUOTATION MARK' }
            0x2026 { 'HORIZONTAL ELLIPSIS' }
            default { 'NON-ASCII CHARACTER' }
        }

        $DisplayValue = switch ($CharacterCode) {
            0x0009 { '<TAB>' }
            0x000A { '<LINE FEED>' }
            0x000D { '<CARRIAGE RETURN>' }
            0x00A0 { '<NON-BREAKING SPACE>' }
            0x200B { '<ZERO-WIDTH SPACE>' }
            default { "'$Character'" }
        }

        $Results.Add("$DisplayValue - $CharacterName ($UnicodeCode)")
    }

    return ($Results -join '; ')
}

function ConvertTo-PrintableAscii {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    if ([string]::IsNullOrEmpty($Value)) {
        return ''
    }

    $ReplacementMap = @{
        0x00A0 = ' '
        0x00AD = '-'
        0x00DF = 'ss'
        0x00C6 = 'AE'
        0x00E6 = 'ae'
        0x00D0 = 'D'
        0x00F0 = 'd'
        0x00D8 = 'O'
        0x00F8 = 'o'
        0x00DE = 'Th'
        0x00FE = 'th'
        0x0110 = 'D'
        0x0111 = 'd'
        0x0141 = 'L'
        0x0142 = 'l'
        0x0152 = 'OE'
        0x0153 = 'oe'
        0x2010 = '-'
        0x2011 = '-'
        0x2012 = '-'
        0x2013 = '-'
        0x2014 = '-'
        0x2015 = '-'
        0x2018 = "'"
        0x2019 = "'"
        0x201A = "'"
        0x201B = "'"
        0x201C = '"'
        0x201D = '"'
        0x201E = '"'
        0x201F = '"'
        0x2022 = '-'
        0x2026 = '...'
        0x2212 = '-'
        0xFF0C = ','
        0xFF0D = '-'
        0xFF0E = '.'
        0xFF5C = '|'
    }

    $Builder = New-Object System.Text.StringBuilder

    foreach ($Character in $Value.ToCharArray()) {
        $CharacterCode = [int]$Character

        if ($CharacterCode -ge 0x20 -and $CharacterCode -le 0x7E) {
            [void]$Builder.Append($Character)
            continue
        }

        if ($ReplacementMap.ContainsKey($CharacterCode)) {
            [void]$Builder.Append([string]$ReplacementMap[$CharacterCode])
            continue
        }

        $DecomposedValue = ([string]$Character).Normalize(
            [System.Text.NormalizationForm]::FormD
        )

        foreach ($DecomposedCharacter in $DecomposedValue.ToCharArray()) {
            $UnicodeCategory = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory(
                $DecomposedCharacter
            )

            if (
                $UnicodeCategory -eq [System.Globalization.UnicodeCategory]::NonSpacingMark -or
                $UnicodeCategory -eq [System.Globalization.UnicodeCategory]::SpacingCombiningMark -or
                $UnicodeCategory -eq [System.Globalization.UnicodeCategory]::EnclosingMark
            ) {
                continue
            }

            $DecomposedCode = [int]$DecomposedCharacter

            if ($DecomposedCode -ge 0x20 -and $DecomposedCode -le 0x7E) {
                [void]$Builder.Append($DecomposedCharacter)
            }
        }
    }

    return $Builder.ToString()
}

function Get-PrimaryContactEmail {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object[]]$Emails
    )

    $AvailableEmails = @(
        $Emails |
            Where-Object {
                $null -ne $_ -and
                -not [string]::IsNullOrWhiteSpace([string]$_.value)
            }
    )

    if ($AvailableEmails.Count -eq 0) {
        return ''
    }

    $PrimaryEmail = $AvailableEmails |
        Where-Object { [bool]$_.primary } |
        Select-Object -First 1

    if ($null -eq $PrimaryEmail) {
        $PrimaryEmail = $AvailableEmails | Select-Object -First 1
    }

    return [string]$PrimaryEmail.value
}

function Test-PhoneNumberStartsWithPlus {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $false
    }

    return [regex]::IsMatch(
        $Value.Trim(),
        $script:PhoneNumberPattern
    )
}

function Get-PhoneDisplayValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Phone
    )

    if ($null -eq $Phone) {
        return ''
    }

    $FormattedValue = [string]$Phone.'formatted-value'

    if (-not [string]::IsNullOrWhiteSpace($FormattedValue)) {
        return $FormattedValue.Trim()
    }

    $RawValue = [string]$Phone.value

    if (-not [string]::IsNullOrWhiteSpace($RawValue)) {
        return $RawValue.Trim()
    }

    return ''
}

function Get-PhonePrefixSuggestion {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ''
    }

    $TrimmedValue = $Value.Trim()

    if ($TrimmedValue.StartsWith('+')) {
        return $TrimmedValue
    }

    # Convert the common international 00 prefix to +.
    if ($TrimmedValue.StartsWith('00') -and $TrimmedValue.Length -gt 2) {
        return '+' + $TrimmedValue.Substring(2)
    }

    return '+' + $TrimmedValue
}

function Write-EmptyAuditCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $Header = '"Organization","ContactName","ContactEmail","IssueType","FieldName","CurrentValue","SuggestedValue","InvalidCharacters","EmailLabel","PrimaryEmail","PhoneLabel","PrimaryPhone","PhoneExtension","PhoneRawValue","ContactURL","ContactId"'
    Set-Content -LiteralPath $Path -Value $Header -Encoding UTF8
}

$Configuration = $null
$Headers = $null

try {
    $Configuration = Show-ConfigurationDialog

    if ($null -eq $Configuration) {
        return
    }

    $BaseUri = switch ([string]$Configuration.Region) {
        'US' { 'https://api.itglue.com' }
        'EU' { 'https://api.eu.itglue.com' }
        'AU' { 'https://api.au.itglue.com' }
        default { throw "Unsupported IT Glue region: $($Configuration.Region)" }
    }

    $Headers = @{
        'x-api-key' = [string]$Configuration.ApiKey
        'Accept'    = 'application/vnd.api+json'
    }

    Write-Host ''
    Write-Host 'Retrieving contacts from IT Glue...'
    Write-Host ''

    $PageSize = 1000
    $PageNumber = 1
    $AllContacts = [System.Collections.Generic.List[object]]::new()

    do {
        $ContactsUri = "$BaseUri/contacts?page%5Bsize%5D=$PageSize&page%5Bnumber%5D=$PageNumber"
        Write-Host "Retrieving page $PageNumber..."

        $Response = Invoke-ITGlueGetRequest -Uri $ContactsUri -Headers $Headers
        $ContactsOnPage = @($Response.data)

        foreach ($Contact in $ContactsOnPage) {
            $AllContacts.Add($Contact)
        }

        $PageNumber++
    }
    while ($ContactsOnPage.Count -eq $PageSize)

    Write-Host ''
    Write-Host "$($AllContacts.Count) contacts retrieved from IT Glue."

    $AuditResults = [System.Collections.Generic.List[object]]::new()

    foreach ($Contact in $AllContacts) {
        $Attributes = $Contact.attributes
        $Emails = @($Attributes.'contact-emails')
        $Phones = @($Attributes.'contact-phones')

        $FirstName = [string]$Attributes.'first-name'
        $LastName = [string]$Attributes.'last-name'
        $Title = [string]$Attributes.title
        $ContactName = [string]$Attributes.name
        $OrganizationName = [string]$Attributes.'organization-name'
        $ContactURL = [string]$Attributes.'resource-url'

        if ([string]::IsNullOrWhiteSpace($ContactName)) {
            $ContactName = ("$FirstName $LastName").Trim()
        }

        $PrimaryContactEmail = Get-PrimaryContactEmail -Emails $Emails

        foreach ($Email in $Emails) {
            if ($null -eq $Email) {
                continue
            }

            $CurrentEmail = [string]$Email.value

            if ([string]::IsNullOrWhiteSpace($CurrentEmail)) {
                continue
            }

            if ([regex]::IsMatch($CurrentEmail, '[A-Z]')) {
                $AuditResults.Add(
                    [PSCustomObject]@{
                        Organization      = $OrganizationName
                        ContactName       = $ContactName
                        ContactEmail      = $CurrentEmail
                        IssueType         = 'Uppercase email address'
                        FieldName         = 'Email'
                        CurrentValue      = $CurrentEmail
                        SuggestedValue    = $CurrentEmail.ToLowerInvariant()
                        InvalidCharacters = ''
                        EmailLabel        = [string]$Email.'label-name'
                        PrimaryEmail      = [bool]$Email.primary
                        PhoneLabel        = ''
                        PrimaryPhone      = ''
                        PhoneExtension    = ''
                        PhoneRawValue     = ''
                        ContactURL        = $ContactURL
                        ContactId         = [string]$Contact.id
                    }
                )
            }
        }

        foreach ($Phone in $Phones) {
            if ($null -eq $Phone) {
                continue
            }

            $CurrentPhone = Get-PhoneDisplayValue -Phone $Phone
            $RawPhone = [string]$Phone.value

            if ([string]::IsNullOrWhiteSpace($CurrentPhone)) {
                continue
            }

            if (-not (Test-PhoneNumberStartsWithPlus -Value $CurrentPhone)) {
                $SuggestedPhone = Get-PhonePrefixSuggestion -Value $CurrentPhone

                $AuditResults.Add(
                    [PSCustomObject]@{
                        Organization      = $OrganizationName
                        ContactName       = $ContactName
                        ContactEmail      = $PrimaryContactEmail
                        IssueType         = 'Phone number missing plus prefix'
                        FieldName         = 'Phone'
                        CurrentValue      = $CurrentPhone
                        SuggestedValue    = $SuggestedPhone
                        InvalidCharacters = ''
                        EmailLabel        = ''
                        PrimaryEmail      = ''
                        PhoneLabel        = [string]$Phone.'label-name'
                        PrimaryPhone      = [bool]$Phone.primary
                        PhoneExtension    = [string]$Phone.extension
                        PhoneRawValue     = $RawPhone
                        ContactURL        = $ContactURL
                        ContactId         = [string]$Contact.id
                    }
                )
            }
        }

        $FieldsToAudit = @(
            [PSCustomObject]@{ FieldName = 'First Name'; Value = $FirstName }
            [PSCustomObject]@{ FieldName = 'Last Name'; Value = $LastName }
            [PSCustomObject]@{ FieldName = 'Title'; Value = $Title }
        )

        foreach ($Field in $FieldsToAudit) {
            if (-not (Test-ContainsInvalidCharacters -Value $Field.Value)) {
                continue
            }

            $AuditResults.Add(
                [PSCustomObject]@{
                    Organization      = $OrganizationName
                    ContactName       = $ContactName
                    ContactEmail      = $PrimaryContactEmail
                    IssueType         = 'Non-ASCII contact field character'
                    FieldName         = [string]$Field.FieldName
                    CurrentValue      = [string]$Field.Value
                    SuggestedValue    = ConvertTo-PrintableAscii -Value ([string]$Field.Value)
                    InvalidCharacters = Get-InvalidCharacterList -Value ([string]$Field.Value)
                    EmailLabel        = ''
                    PrimaryEmail      = ''
                    PhoneLabel        = ''
                    PrimaryPhone      = ''
                    PhoneExtension    = ''
                    PhoneRawValue     = ''
                    ContactURL        = $ContactURL
                    ContactId         = [string]$Contact.id
                }
            )
        }
    }

    $SortedResults = @(
        $AuditResults |
            Sort-Object Organization, ContactName, IssueType, FieldName, CurrentValue
    )

    if ($SortedResults.Count -eq 0) {
        Write-EmptyAuditCsv -Path ([string]$Configuration.ReportPath)

        [System.Windows.Forms.MessageBox]::Show(
            (
                "The audit completed successfully.`r`n`r`n" +
                "No uppercase email addresses, non-ASCII contact field characters, or phone numbers missing a plus prefix were found.`r`n`r`n" +
                "Report:`r`n$($Configuration.ReportPath)"
            ),
            'IT Glue Audit Completed',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    else {
        $SortedResults |
            Export-Csv `
                -LiteralPath ([string]$Configuration.ReportPath) `
                -NoTypeInformation `
                -Encoding UTF8

        $EmailIssueCount = @(
            $SortedResults |
                Where-Object { $_.IssueType -eq 'Uppercase email address' }
        ).Count

        $InvalidFieldIssueCount = @(
            $SortedResults |
                Where-Object { $_.IssueType -eq 'Non-ASCII contact field character' }
        ).Count

        $PhoneIssueCount = @(
            $SortedResults |
                Where-Object { $_.IssueType -eq 'Phone number missing plus prefix' }
        ).Count

        $AffectedContactCount = @(
            $SortedResults.ContactId | Sort-Object -Unique
        ).Count

        Write-Host ''
        Write-Host "$($SortedResults.Count) total issues were found."
        Write-Host "Affected contacts: $AffectedContactCount"
        Write-Host "Uppercase email addresses: $EmailIssueCount"
        Write-Host "Non-ASCII contact field characters: $InvalidFieldIssueCount"
        Write-Host "Phone numbers missing plus prefix: $PhoneIssueCount"
        Write-Host ''
        Write-Host "Report created: $($Configuration.ReportPath)"
        Write-Host ''

        $SortedResults |
            Format-Table Organization, ContactName, ContactEmail, IssueType, FieldName, CurrentValue, SuggestedValue, InvalidCharacters, ContactURL -AutoSize

        [System.Windows.Forms.MessageBox]::Show(
            (
                "The audit completed successfully.`r`n`r`n" +
                "Total issues: $($SortedResults.Count)`r`n" +
                "Affected contacts: $AffectedContactCount`r`n" +
                "Uppercase email addresses: $EmailIssueCount`r`n" +
                "Non-ASCII contact field characters: $InvalidFieldIssueCount`r`n" +
                "Phone numbers missing plus prefix: $PhoneIssueCount`r`n`r`n" +
                "Report:`r`n$($Configuration.ReportPath)"
            ),
            'IT Glue Audit Completed',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }

    if (
        [bool]$Configuration.OpenReport -and
        (Test-Path -LiteralPath ([string]$Configuration.ReportPath) -PathType Leaf)
    ) {
        Invoke-Item -LiteralPath ([string]$Configuration.ReportPath)
    }
}
catch {
    $ErrorMessage = $_.Exception.Message

    if (
        $null -ne $_.ErrorDetails -and
        -not [string]::IsNullOrWhiteSpace($_.ErrorDetails.Message)
    ) {
        $ErrorMessage = $_.ErrorDetails.Message
    }

    [System.Windows.Forms.MessageBox]::Show(
        "The IT Glue audit could not be completed.`r`n`r`n$ErrorMessage",
        'IT Glue Audit Error',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null

    Write-Error -Message $ErrorMessage -ErrorAction Continue
}
finally {
    if ($null -ne $Configuration) {
        foreach ($ConfigurationItem in @($Configuration)) {
            if (
                $null -ne $ConfigurationItem -and
                $null -ne $ConfigurationItem.PSObject.Properties['ApiKey']
            ) {
                $ConfigurationItem.ApiKey = $null
            }
        }
    }

    $Headers = $null
}