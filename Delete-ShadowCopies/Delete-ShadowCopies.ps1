# List all shadow copies
$vssList = vssadmin list shadows

# Filter out the newest restore point and delete the rest
$shadowCopies = $vssList | Where-Object {$_ -match "Shadow Copy ID:"}
$shadowIds = $shadowCopies | ForEach-Object { $_.Split(":")[1].Trim() }

# Keep only the newest restore point (assuming it's the first in the list)
$keepId = $shadowIds[0]

# Delete all other restore points
foreach ($id in $shadowIds | Where-Object {$_ -ne $keepId}) {
    vssadmin delete shadows /shadow=$id /quiet
}