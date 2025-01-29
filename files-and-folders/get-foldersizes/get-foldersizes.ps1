Get-ChildItem -Path 'C:\' -Directory -Force | ForEach-Object {
    $size = (Get-ChildItem $_.FullName -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
    [PSCustomObject]@{
        FolderName = $_.FullName
        SizeGB = [Math]::Round($size / 1GB, 2)
        IsHidden = $_.Attributes.HasFlag([System.IO.FileAttributes]::Hidden)
    }
} | Sort-Object SizeGB -Descending | Select-Object -First 10
