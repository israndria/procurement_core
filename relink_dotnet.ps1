param([string]$ExcelPath)

[System.Reflection.Assembly]::LoadWithPartialName('System.IO.Compression') | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem') | Out-Null

$folder = Split-Path $ExcelPath
$excelName = Split-Path $ExcelPath -Leaf

# URL-encode path for file:/// URI
$excelUri = "file:///" + $ExcelPath.Replace('\', '/').Replace(' ', '%20').Replace('@', '%40')

# Mapping
$mapping = @{
    "BA PK"  = "satu_data"
    "Reviu"  = "list_reviu"
    "Dokpil" = "list_dokpil"
}

$results = @()

Get-ChildItem $folder -Filter "*.docx" | Where-Object {
    $_.Name -notlike "*(Merged)*" -and $_.Name -notlike "~*"
} | ForEach-Object {
    $docxPath = $_.FullName
    $docxName = $_.Name

    # Detect sheet
    $sheet = $null
    foreach ($kw in $mapping.Keys) {
        if ($docxName -match [regex]::Escape($kw)) {
            $sheet = $mapping[$kw]
            break
        }
    }

    if (-not $sheet) {
        $results += "[SKIP] $docxName"
        return
    }

    try {
        $zip = [System.IO.Compression.ZipFile]::Open($docxPath, 'Update')

        # 1. Update settings.xml - add/replace mailMerge
        $entry = $zip.GetEntry('word/settings.xml')
        $stream = $entry.Open()
        $reader = New-Object System.IO.StreamReader($stream)
        $xml = $reader.ReadToEnd()
        $reader.Close()
        $stream.Close()

        # Remove existing mailMerge
        $xml = [regex]::Replace($xml, '<w:mailMerge>.*?</w:mailMerge>', '', 'Singleline')

        # Build new mailMerge
        $escaped = $ExcelPath.Replace('&', '&amp;')
        $mailMerge = "<w:mailMerge>" +
            '<w:mainDocumentType w:val="formLetters"/>' +
            '<w:linkToQuery/>' +
            '<w:dataType w:val="native"/>' +
            "<w:connectString w:val=`"Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=$escaped;Mode=Read;Extended Properties=&quot;HDR=YES;IMEX=1&quot;;`"/>" +
            "<w:query w:val=`"SELECT * FROM ``$sheet`$```"/>" +
            "</w:mailMerge>"

        $xml = $xml.Replace('</w:settings>', $mailMerge + '</w:settings>')

        # Write back - truncate and rewrite
        $entry.Delete()
        $newEntry = $zip.CreateEntry('word/settings.xml', 'Optimal')
        $stream = $newEntry.Open()
        $writer = New-Object System.IO.StreamWriter($stream, [System.Text.Encoding]::UTF8)
        $writer.Write($xml)
        $writer.Close()
        $stream.Close()

        # 2. Update/create settings.xml.rels
        $relsEntry = $zip.GetEntry('word/_rels/settings.xml.rels')
        if ($relsEntry) { $relsEntry.Delete() }

        $relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            "<Relationship Id=`"rId1`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeSource`" Target=`"$excelUri`" TargetMode=`"External`"/>" +
            "<Relationship Id=`"rId2`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeSource`" Target=`"$excelUri`" TargetMode=`"External`"/>" +
            '</Relationships>'

        $newRels = $zip.CreateEntry('word/_rels/settings.xml.rels', 'Optimal')
        $stream = $newRels.Open()
        $writer = New-Object System.IO.StreamWriter($stream, [System.Text.Encoding]::UTF8)
        $writer.Write($relsXml)
        $writer.Close()
        $stream.Close()

        $zip.Dispose()
        $results += "[OK] $docxName -> $sheet"
    } catch {
        if ($zip) { $zip.Dispose() }
        $results += "[GAGAL] $docxName ($_)"
    }
}

# Show result popup
Add-Type -AssemblyName System.Windows.Forms
$msg = "Relink selesai!`n`n" + ($results -join "`n")
[System.Windows.Forms.MessageBox]::Show($msg, "Relink Template", "OK", "Information") | Out-Null
