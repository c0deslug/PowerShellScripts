@FailedDevice = @()

$ContentPath = "C:\Users\<<username>>\Desktop\DeviceList.txt"

$TargetCollectionName = "Collection Name"

Get-Content $ContentPath | ForEach-Object {
  $item = $_
  try{
  Write-Host "Adding $($_) to collection..." -ForegroundColor Green
  Add-CMDeviceCollectionMembershipRule -CollectionName $TargetCollectionName -ResourceID (Get-CMDevice -Name $_).ResourceID
  }catch{
  Write-Host "Failed to add $item to collection, adding it to FailedDevices variable..." -ForegroundColor Red
  Write-Host "$_"
  $FailedDevices += $item
  }
}
Write-Host "Finished, check for failed devices by accessing the `$FailedDevices variable" -ForegroundColor Yellow
      
