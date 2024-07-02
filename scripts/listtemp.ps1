# Script to list contents of directory C:\temp

# Specify the directory path
$directory = "C:\temp"

# Check if the directory exists
if (Test-Path $directory -PathType Container) {
    # Get the list of items in the directory
    $items = Get-ChildItem $directory

    # Output each item in the directory
    foreach ($item in $items) {
        Write-Output $item.FullName
    }
} else {
    Write-Output "Directory $directory does not exist."
}
