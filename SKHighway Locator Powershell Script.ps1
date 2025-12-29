
<# 
    Photo CSV + KML generator (recursive, relative paths, TopMost folder picker with "new UI")
    - Select a root folder (dialog appears in front of ISE; you can paste a path or browse)
    - Scans JPG/JPEG in folder and all subfolders
    - Writes photo_metadata.csv and photo_map.kml in the root folder
    - CSV includes Excel hyperlink formulas with relative paths + PhotoFolder column (relative directory)
    - KML placemarks include an <img> using a relative src plus a text description block
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# -----------------------------
# Option A: Shell COM folder picker (editable path) with TopMost owner
# -----------------------------
function Select-RootFolder {
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $selected = $null

    # Tiny TopMost owner so dialog stays in front of ISE
    $owner = New-Object System.Windows.Forms.Form
    $owner.TopMost = $true
    $owner.StartPosition = 'CenterScreen'
    $owner.ShowInTaskbar = $false
    $owner.FormBorderStyle = 'None'
    $owner.Size = [System.Drawing.Size]::new(1,1)
    $owner.Show()

    try {
        # Flags: BIF_USENEWUI (0x50) + BIF_RETURNONLYFSDIRS (0x01) = 0x51
        $flags = 0x51
        $shell = New-Object -ComObject Shell.Application
        $folder = $shell.BrowseForFolder($owner.Handle, "Select a folder (you can paste a path)", $flags)
        if ($folder -and $folder.Self) {
            $selected = $folder.Self.Path
        }
    } finally {
        $owner.Dispose()
    }

    return $selected
}

# -----------------------------
# Select root folder
# -----------------------------
$selectedFolder = Select-RootFolder
if (-not $selectedFolder -or -not (Test-Path $selectedFolder)) {
    Write-Host "No folder selected. Exiting."
    exit
}

$folderName = Split-Path $selectedFolder -Leaf
$csvPath = Join-Path $selectedFolder "photo_metadata.csv"
$kmlPath = Join-Path $selectedFolder "photo_map.kml"

# -----------------------------
# Get metadata via Shell COM
# -----------------------------
function Get-FileMetadata {
    param ([string]$filePath)

    $shell = New-Object -ComObject Shell.Application
    $folder = $shell.Namespace((Split-Path $filePath))
    $file = $folder.ParseName((Split-Path $filePath -Leaf))

    $metadata = @{}
    for ($i = 0; $i -lt 300; $i++) {
        $key = $folder.GetDetailsOf($folder.Items, $i)
        $value = $folder.GetDetailsOf($file, $i)
        if ($key -and $value) { $metadata[$key] = $value }
    }
    return $metadata
}

# -----------------------------
# Read GPS from EXIF (decimal degrees)
# -----------------------------
function Get-GpsCoordinates {
    param ([string]$imagePath)
    $img = $null
    try {
        $img = [System.Drawing.Image]::FromFile($imagePath)
        $props = $img.PropertyItems
        $gpsLatRef = ($props | Where-Object { $_.Id -eq 1 }).Value
        $gpsLat    = ($props | Where-Object { $_.Id -eq 2 }).Value
        $gpsLonRef = ($props | Where-Object { $_.Id -eq 3 }).Value
        $gpsLon    = ($props | Where-Object { $_.Id -eq 4 }).Value

        if (-not $gpsLat -or -not $gpsLon) { return $null }

        function DecodeCoordinate {
            param([byte[]]$bytes)
            $rational = @()
            for ($i = 0; $i -lt $bytes.Length; $i += 8) {
                $num = [BitConverter]::ToUInt32($bytes, $i)
                $den = [BitConverter]::ToUInt32($bytes, $i + 4)
                if ($den -ne 0) { $rational += ($num / $den) }
            }
            return $rational[0] + ($rational[1] / 60) + ($rational[2] / 3600)
        }

        $lat = DecodeCoordinate $gpsLat
        $lon = DecodeCoordinate $gpsLon
        $latRef = [System.Text.Encoding]::ASCII.GetString($gpsLatRef).Trim([char]0)
        $lonRef = [System.Text.Encoding]::ASCII.GetString($gpsLonRef).Trim([char]0)

        if ($latRef -eq "S") { $lat = -$lat }
        if ($lonRef -eq "W") { $lon = -$lon }

        return @{ Latitude = $lat; Longitude = $lon }
    } catch {
        return $null
    } finally {
        if ($img) { $img.Dispose() }
    }
}

# -----------------------------
# Parse labeled fields from Title
# -----------------------------
function Parse-TitleFields {
    param ([string]$title)

    $fields = @{
        Road     = ""
        Station  = ""
        Offset   = ""
        Accuracy = ""
        Note     = ""
    }

    if ([string]::IsNullOrWhiteSpace($title)) { return $fields }

    $segments = ($title -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    foreach ($seg in $segments) {
        if ($seg -match '^\s*(?<key>Road|Station|Offset|Accuracy|Note)\s*:\s*(?<val>.+?)\s*$') {
            $fields[$matches['key']] = $matches['val'].Trim()
        }
    }
    return $fields
}

# -----------------------------
# Compute relative path (selectedFolder -> file)
# -----------------------------
function Get-RelativePath {
    param ([string]$baseFolder, [string]$fullPath)
    # Ensure base is treated as a folder by Uri
    $baseWithSep = if ($baseFolder[-1] -eq [IO.Path]::DirectorySeparatorChar) { $baseFolder } else { $baseFolder + [IO.Path]::DirectorySeparatorChar }
    $uBase = New-Object System.Uri($baseWithSep)
    $uFull = New-Object System.Uri($fullPath)
    $rel = $uBase.MakeRelativeUri($uFull).ToString()         # forward slashes
    [System.Uri]::UnescapeDataString($rel)
}

# -----------------------------
# Collect photo metadata (recursive)
# -----------------------------
$photoData = @()

$photos = Get-ChildItem -Path $selectedFolder -File -Recurse | Where-Object {
    $_.Extension -match '\.jpe?g$'
}

foreach ($file in $photos) {
    $metadata = Get-FileMetadata $file.FullName
    $coords = Get-GpsCoordinates $file.FullName

    $title = $metadata["Title"]
    $subject = $metadata["Subject"]

    # Labeled fields from Title
    $labelFields = Parse-TitleFields -title $title

    # Date Taken
    $dateTaken = ""
    foreach ($key in $metadata.Keys) {
        if ($key -match "Date\s*taken") { $dateTaken = $metadata[$key]; break }
    }
    if (-not $dateTaken -and $metadata["Date created"]) { $dateTaken = $metadata["Date created"] }

    # Coordinates
    $latitude  = if ($coords) { $coords["Latitude"] }  else { $null }
    $longitude = if ($coords) { $coords["Longitude"] } else { $null }

    # Relative paths (Windows for Excel; URL-style for KML)
    $relUrl     = Get-RelativePath -baseFolder $selectedFolder -fullPath $file.FullName   # e.g., Sub/IMG_0001.jpg
    $relForExcel = ($relUrl -replace '/', '\')                                            # backslashes for Excel on Windows
    $relForKml   = ($relUrl -replace '\\', '/')                                           # forward slashes for KML/URLs

    # Photo folder (relative directory). If photo is in root, show "."
    $photoFolderRel = Split-Path -Path $relForExcel -Parent
    if ([string]::IsNullOrWhiteSpace($photoFolderRel)) { $photoFolderRel = "." }

    # Excel hyperlink formula (no backticks): =HYPERLINK("Rel\Path.jpg","Open Photo")
    $photoLinkFormula = ('=HYPERLINK("{0}","Open Photo")' -f $relForExcel)

    # Build object
    $obj = [ordered]@{
        FileName    = $file.Name
        PhotoFolder = $photoFolderRel
        Subject     = $subject
        DateTaken   = $dateTaken
        Latitude    = $latitude
        Longitude   = $longitude

        Road        = $labelFields["Road"]
        Station     = $labelFields["Station"]
        Offset      = $labelFields["Offset"]
        Accuracy    = $labelFields["Accuracy"]
        Note        = $labelFields["Note"]

        PhotoLink   = $photoLinkFormula
        KmlSrcPath  = $relForKml   # internal use when building KML
    }

    $photoData += New-Object PSObject -Property $obj
}

# -----------------------------
# Export CSV
# -----------------------------
$photoData |
    Select-Object FileName,PhotoFolder,Subject,DateTaken,Latitude,Longitude,Road,Station,Offset,Accuracy,Note,PhotoLink |
    Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host ("CSV file created at: {0}" -f $csvPath)

# -----------------------------
# Build KML (image + text inside CDATA, relative src)
# -----------------------------
$kml = @(
'<?xml version="1.0" encoding="UTF-8"?>',
'<kml xmlns="http://earth.google.com/kml/2.0">',
'  <Document>',
'    <name>My Photos</name>',
'    <open>1</open>',
'    <Style id="Photo">',
'      <IconStyle>',
'        <Icon>',
'          <href>http://maps.google.com/mapfiles/kml/pal4/icon38.png</href>',
'          <scale>1.0</scale>',
'        </Icon>',
'      </IconStyle>',
'    </Style>',
'    <Folder>',
("      <name>{0}</name>" -f $folderName),
'      <open>0</open>'
)

foreach ($p in $photoData) {
    if ($p.Latitude -and $p.Longitude) {
        # Placemark name preference: Note > Road+Station > Road > FileName
        $placemarkName =
            if ($p.Note) {
                $p.Note
            } elseif ($p.Road -and $p.Station) {
                ("{0} – Station {1}" -f $p.Road, $p.Station)
            } elseif ($p.Road) {
                $p.Road
            } else {
                $p.FileName
            }

        # Description text block (use Environment.NewLine for reliability)
        $nl = [Environment]::NewLine
        $desc = (
            "File: {0}{7}Folder: {1}{7}Road: {2}{7}Station: {3}{7}Offset: {4}{7}Accuracy: {5}{7}Note: {6}" `
            -f $p.FileName, $p.PhotoFolder, $p.Road, $p.Station, $p.Offset, $p.Accuracy, $p.Note, $nl
        )

        # Build the <img> tag with relative src then append text block
        $imgTag   = ('<img src="{0}" style="max-width:500px;max-height:500px;">' -f $p.KmlSrcPath)
        $descHtml = $imgTag + ('<div style="white-space:pre-line;font-family:sans-serif;font-size:12px;margin-top:6px;">{0}</div>' -f $desc)

        $kml += @(
'      <Placemark>',
("        <description><![CDATA[{0}]]></description>" -f $descHtml),
'        <Snippet/>',
("        <name>{0}</name>" -f $placemarkName),
'        <styleUrl>#Photo</styleUrl>',
'        <Point>',
'          <altitudeMode>clampedToGround</altitudeMode>',
("          <coordinates>{0},{1},0</coordinates>" -f $p.Longitude, $p.Latitude),
'        </Point>',
'      </Placemark>'
        )
    }
}

$kml += @(
'    </Folder>',
'  </Document>',
'</kml>'
)

$kml | Out-File -FilePath $kmlPath -Encoding UTF8
Write-Host ("KML file created at: {0}" -f $kmlPath)
