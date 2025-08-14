function CreateAndPopulateDeviceCollections {
    param (
        [Parameter(Mandatory = $true, HelpMessage = "Name of the collections you want to create, they will use this pattern [ Name - Deployment Batch <number>]")][string]$CollectionName,
        [Parameter(Mandatory = $true, HelpMessage = "Name of the limiting collection for the collections that will be created")][string]$LimitingCollectionName,
        [Parameter(Mandatory = $true, HelpMessage = "Enter the number of collections you want to create")][int]$NumberOfCollections,
        [Parameter(Mandatory = $true, HelpMessage = "Comment for created collection, ticket number if applicable, initials and date of creation")][string]$Comment,
        [Parameter(Mandatory = $true, HelpMessage = "Enter the folder path on where to move the collections, check Object Path in SCCM console for help")][string]$FolderPath,
        [Parameter(Mandatory = $true, HelpMessage = "Enter Site code")][string]$SiteCode,
        [Parameter(Mandatory = $true, HelpMessage = "Emter the absolute file path to the CSV; the CSV must contain the coloumns 'Deployment Batch' and 'Device Name'; The Deployment Batch column values should contain a number within the string that will allow for the function to select and add  only the devices inteded for the collection with the same deployment batch number; For example devices that have the number 5 in the string in the Deployment Batch column will be added to the <Name - Deployment Batch 5> collection ; The delimiter must be a semicolon!")][string]$CSVPath
    )


$data = Import-Csv -Path $CSVPath -Delimiter ';'

for ($i = 1; $i -le $NumberOfCollections; i++) {
    $NewCollectionName = "$CollectionName - Deployment Batch $i"
    $Null = New-CMCollection -CollectionType Device -Comment $Comment -LimitingCollectionName $LimitingCollectionName -Name $NewCollectionName
    $InputCollection = Get-CMCollection -Name $NewCollectionName
    Write-Host "Created $NewCollectionName collection in the root folder." -ForegroundColor Yellow
    Move-CMObject -InputObject $InputCollection -FolderPath "${SiteCode}:\DeviceCollection\$FolderPath"
    Write-Host "Moved $NewCollectionName collection to $FolderPath folder." -ForegroundColor Yellow
    $FR_Array = $data | Where-Object { $_.'Deployment Batch' -match "\b$i\b" } | Select-Object -ExpandProperty 'Device Name'
    foreach ($device in $FR_Array) {
        try{
        Add-CMDeviceCollectionDirectMemebershipRule -CollectionName $NewCollectionName -ResourceId (Get-CMDevice -Name $device).ResourceId
        Write-Host "Added $device to $NewCollectionName collection." -ForegroundColor Green
        } catch {
            Write-Host "Failed to add $device to $NewCollectionName collection. Error: $_" -ForegroundColor Red
        }
    }
}

}
