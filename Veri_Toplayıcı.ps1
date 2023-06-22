#excel import code:  Install-Module -Name ImportExcel -Scope CurrentUser


# Import the module
Import-Module -Name ImportExcel

# Specify the folder path containing the text files
$folderPath = "Type here your folder's filepath"  

# Specify the Excel file path
$excelFilePath = "$env:USERPROFILE\Desktop\excelfile.xlsx"

# Initialize an empty array to store all the data from text files
$allData = @()
$diskBoyutuList = @()

# Get all the text files in the folder
$textFiles = Get-ChildItem -Path $folderPath -Filter "*.txt"

# Process each text file
foreach ($file in $textFiles) {
    # Read the content of the text file
    $content = Get-Content -Path $file.FullName

    # Extract the desired elements from the text
    $name = ($content | Select-String -Pattern "^Isim: (.+)$").Matches.Groups[1].Value
    $firma = ($content | Select-String -Pattern "^Firma: (.+)$").Matches.Groups[1].Value
    $model = ($content | Select-String -Pattern "^Model: (.+)$").Matches.Groups[1].Value
    $seriNumarasi = ($content | Select-String -Pattern "^Seri Numarasi: (.+)$").Matches.Groups[1].Value
    $isletimsistemi = ($content | Select-String -Pattern "^Isletim Sistemi: (.+)$").Matches.Groups[1].Value
    $islemci = ($content | Select-String -Pattern "^Islemci: (.+)$").Matches.Groups[1].Value
    $ekrankarti = ($content | Select-String -Pattern "^Ekran Karti: (.+)$").Matches.Groups[1].Value
    $ramboyutu = ($content | Select-String -Pattern "^Toplam RAM boyutu: (.+)$").Matches.Groups[1].Value
    $ramboyutu = ($content | Select-String -Pattern "^Ram miktari: (.+)$").Matches.Groups[1].Value
    $macadress = ($content | Select-String -Pattern "^Physical MAC Address: (.+)$").Matches.Groups[1].Value
    $macadress = ($content | Select-String -Pattern "^mac: (.+)$").Matches.Groups[1].Value

    $diskBoyutuLines = $content | Select-String -Pattern "^Disk boyutu: (.+)$" | ForEach-Object { $_.Matches.Groups[1].Value }

    $diskBoyutuList += $diskBoyutuLines

    # Calculate the total disk boyutu
    $totalDiskBoyutu = 0
    foreach ($number in $diskBoyutuList) {
        $number = $number.Substring(0, $number.Length - 2)
        $number = [int]$number

        if ($number -gt 64) {
            $totalDiskBoyutu += $number
        }
    }


    $diskType = if ($content -match "SSD") { "SSD" } else { "SATA" }

    $diskBoyutuList = @()
    $fileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
    $fileNameWithoutExtension = $fileNameWithoutExtension.Replace('.', '/')

    # Create a custom object with the extracted data, including the file name and selected "Disk boyutu" value
    $dataObject = [PSCustomObject]@{
        SICIL_NUMARASI = $fileNameWithoutExtension
        ISIM = $name
        FIRMA = $firma
        MODEL = $model
        SERI_NUMARASI = $seriNumarasi
        ISLETIM_SISTEMI = $isletimsistemi
        ISLEMCI = $islemci
        EKRAN_KARTI = $ekrankarti
        RAM_BOYUTU = $ramboyutu
        MAC_ADRESI = $macadress
        DISK_BOYUTU = $totalDiskBoyutu
        DISK_TIPI = $diskType
    }

    # Add the data object to the array
    $allData += $dataObject
}

# Write all the data to the Excel file
$allData | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter
