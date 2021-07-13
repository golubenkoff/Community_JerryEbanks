Function Get-AdSitesAndSubnets 
{
##  Print a list of site and their subnets

$configNCDN = (Get-ADRootDSE).ConfigurationNamingContext

$siteContainerDN = ("CN=Sites," + $configNCDN)

$siteObjs = Get-ADObject -SearchBase $siteContainerDN –filter { objectClass -eq "site" } -properties "siteObjectBL", name
[array]$Sites = @()
foreach ($siteObj in $siteObjs) {

    $subnetArray = New-Object -Type string[] -ArgumentList $siteObj.siteObjectBL.Count

    $i = 0

    foreach ($subnetDN in $siteObj.siteObjectBL) {

        $subnetName = $subnetDN.SubString(3, $subnetDN.IndexOf(",CN=Subnets,CN=Sites,") – 3)

        $subnetArray[$i] = $subnetName

        $i++

    }

    $siteSubnetObj = New-Object PSCustomObject | Select SiteName, Subnets

    $siteSubnetObj.SiteName = $siteObj.Name

    $siteSubnetObj.Subnets = $subnetArray

    $siteSubnetObj
    $Sites += $siteSubnetObj
}

return $Sites
}


Get-AdSitesAndSubnets 