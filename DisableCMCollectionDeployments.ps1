$TargetCollectionName = "Collection Name"
$deployments = Get-CMApplicationDeployment -CollectionName $TargetCollectionName
$deployments.forEach{$_.Enabled = $false; $_.put()} 

# Note: You can also use ApplicationName instead of CollectionName
