$OutputEncoding = [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding

#ここのパスを変更
$rootDir = "pass"

$epubFiles = Get-ChildItem -LiteralPath $rootDir -Filter "*.epub" | Sort-Object Name

$volumeNumber = 1

foreach ($epub in $epubFiles) {
    $epubPath = $epub.FullName
    $zipPath = $epubPath -replace "\.epub$", ".zip"

    Rename-Item -LiteralPath $epubPath -NewName $zipPath -Force

    $extractDir = Join-Path $rootDir ([System.IO.Path]::GetFileNameWithoutExtension($epub.Name))

    Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force

    $imageFolder = Join-Path $extractDir "OEBPS\image"

    if (Test-Path $imageFolder) {
        $newImageFolder = Join-Path $rootDir "$volumeNumber"

        if (Test-Path $newImageFolder) {
            Remove-Item -LiteralPath $newImageFolder -Recurse -Force
        }

        Move-Item -LiteralPath $imageFolder -Destination $newImageFolder

        $volumeNumber++
    }

    Remove-Item -LiteralPath $zipPath -Force
}
