#    Save it as e.g D:\StartScan.ps1
#    Create a new shortcut and point it to
#    %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe -File "D:\StartScan.ps1"
#    Change Item("FormatID").Value = $wiaFormatJPEG to $wiaFormatPNG (or TIFF, BMP, GIF) if you need another image format
#    Change $([Environment]::GetFolderPath("Desktop"))\Scan {0}.jpg" if you need another output path. 
#	Change the extension .jpg if you previously had changed the image format
#    open scan folder
#    open latest document in folder


# Create object to access the scanner
$deviceManager = new-object -ComObject WIA.DeviceManager
$device = $deviceManager.DeviceInfos.Item(1).Connect()

# Create object to access the scanned image later
$imageProcess = new-object -ComObject WIA.ImageProcess

# Store file format GUID strings
$wiaFormatBMP  = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatPNG  = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatGIF  = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatJPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatTIFF = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"

# Scan the image from scanner as BMP
foreach ($item in $device.Items) {
    $image = $item.Transfer() 
}

# set type to TIFF and quality/compression level
$imageProcess.Filters.Add($imageProcess.FilterInfos.Item("Convert").FilterID)
$imageProcess.Filters.Item(1).Properties.Item("FormatID").Value = $wiaFormatTIFF
$imageProcess.Filters.Item(1).Properties.Item("Quality").Value = 5
$image = $imageProcess.Apply($image)

# Build filepath from desktop path and filename 'Scan 0'
$filename = "$([Environment]::GetFolderPath("P:\Scans"))\Scan {0}.tiff"

# If a file named 'Scan 0' already exists, increment the index as long as needed
$index = 0
while (test-path ($filename -f $index)) {[void](++$index)}
$filename = $filename -f $index

# Save image to '\\hubcloud\Public\Scans {x}'
$image.SaveFile($filename)

# Show image 
& $filename

$objShell = New-Object -ComObject "Shell.Application"
$objShell.Explore("P:\Scans")