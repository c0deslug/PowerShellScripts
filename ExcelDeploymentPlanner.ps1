

Install-Module ImportExcel -Force -Scope CurrentUser
Import-Module ImportExcel

function Get-NextWeekday{
    param (
        [datetime]$startDate,
        [int[]]$weekdays # Expected to be @(2, 3, 4) for Tue, Wed, Thu
    )
    for ($i = 0; $i -lt $i++) {
        
        $nextDate = $startDate.AddDays($i)
        
        if ($nextDate.DayOfWeek.Value__ -in $weekdays) {
            return $nextDate
        }

    }
    return $null #This shoul never be reached
}

function Get-DeploymentDates {
    param (
        [datetime]$startDate,
        [int]$totalBatches
    )
    $weekdays = @(2, 3, 4) # Tuesday, Wednesday, Thursday
    $deploymentDates = @()
    $currentDate = $startDate
    while ($deploymentDates.Count -lt $totalBatches) {
        # Get the next eligible weekday date
        $nextDate = Get-NextWeekday -startDate $currentDate -weekdays $weekdays

        # Add the result to the deployment dates and move to the next date
        $deploymentDates += $nextDate
        $currentDate = $nextDate.AddDays(1) 
    }
    return $deploymentDates
}

function Assign-BatchesAndDates {
    param (
        [array]$data,
        [datetime]$startDate
    )
    
    $batchSizes = @(1, 3, 5, 10, 15, 30, 60, 90, 120, 150, 150, 150, 300)
    # Group by Country Code and apply batch assignments

    $groupedData = $data | Group-Object -Property "Country Code"
    foreach ($group in $groupedData) {
        $devices = $group.Group
        $currentIndex = 0
        $totalDevices = $devices.Count
        $batchNumber = 1
        $startIndex = 0

        while ($totalDevices -gt 0) {
            if ($currentIndex -lt (batchSizes.Count)){
                $batchSize = $batchSizes[$currentIndex]}
                else{ $batchSize = batchSizes[-1] } # Use the last batch size if we exceed the list
            $endIndex = [math]::Min($startIndex + $batchSize, $devices.Count)
            # Assign batch and deployment date for each device in the batch
            for ($i = $startIndex; $i -lt $endIndex; $i++) {
                $devices[$i]."Deployment Batch" = "Batch $batchNumber"
            }
            $deploymentDates = Get-DeploymentDates -startDate $startDate -totalBatches $batchNumber
            $deploymentDate = $deploymentDates[$batchNumber - 1]
            for ($i = $startIndex; $i -lt $endIndex; $i++) {
                $devices[$i]."Deployment Dates" = $deploymentDate.ToString("yyyy-MM-dd")
            }
            $batchNumber++
            $totalDevices -= ($endIndex - $startIndex)
            $startIndex = $endIndex
            $currentIndex++
        }
    }
    return $data
}
#Load Excel file
$filePath = Read-Host "Enter the path to the Excel file"
$outputFilePath = Read-Host "Enter the path to save the output Excel file"
#Set the start date for deployments
Write-Host "Deployment dates will be generated for every Tuesday, Wednesday, and Thursday starting with the given start date in the format of YYYY-MM-DD."
$startDateStr = Read-Host "Enter the start date (YYYY-MM-DD)"
$startDate = [datetime]::ParseExact($startDateStr, "yyyy-MM-dd", $null)
#Load data into an array of objects
$data = Import-Excel -Path $filePath
$requiredColumns = @("Country Code", "Device Name", "Deployment Batch", "Deployment Dates")
foreach ($col in $requiredColumns) {
    if (-not ($data | Select-Object -First 1 | Get-Member -Name $col)) {
        throw "The provided Excel file does not contain the required '$col' column." 
        return
    }
}
# Apply batch assignment
$proocessedData = Assign-BatchesAndDates -data $data -startDate $startDate
# Export the modified data to a new Excel file
$proocessedData | Export-Excel -Path $outputFilePath #-NoNumberConversion
Write-Host "Processed data saved to $outputFilePath"
