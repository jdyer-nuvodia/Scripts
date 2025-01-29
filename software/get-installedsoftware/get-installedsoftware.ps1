$InstalledSoftware = Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall"

$results = foreach($obj in $InstalledSoftware) {
  $name = $obj.GetValue('DisplayName')
  $version = $obj.GetValue('DisplayVersion')
  if ($name) {
    [PSCustomObject]@{
      Name = $name
      Version = $version
    }
  }
}

$results | Export-Csv -Path "C:\InstalledSoftware.csv" -NoTypeInformation