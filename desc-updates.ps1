Import-Module ActiveDirectory

# Set the path variable to the location of the script
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Create a subdirectory in the script path called 'Completed Update CSVs' if it doesn't exist
if (! (Test-Path -Path "$scriptPath\Completed Update CSVs")) {
    New-Item -Path "$scriptPath\Completed Update CSVs" -ItemType Directory -Force
}

# Set a variable to the Completed Update CSVs directory to put the completed CSVs in after the operation is complete
$completedPath = "$scriptPath\Completed Update CSVs"

# Make a foreach loop that goes through each file in the directory ending in .csv
foreach ($file in Get-ChildItem -Path "$scriptPath" -Filter "*.csv") {

    $spreadsheetData = Import-Csv -Path $file.FullName

    # Go through each row in the spreadsheet data
    foreach ($row in $spreadsheetData) {
        # Get the CN from $row.AD_CN
        $cn = $row.AD_CN

        # Modify the description of $cn using $row.NewDescription
        if ($cn -and $row.NewDescription) {
            Set-ADGroup -Identity $cn -Description $row.NewDescription
        }

    }

    # Move the file to the Completed Update CSVs directory
    Move-Item -Path $file.FullName -Destination $completedPath

}

