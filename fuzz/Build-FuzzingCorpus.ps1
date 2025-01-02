
function Build-FuzzingCorpus {
    param (
        [string]$InputDir,
        [string]$OutputDir
    )

    # Ensure the output directory exists
    if (-not (Test-Path -Path $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir
    }

    # Function to convert hex string to byte array
    function Convert-HexStringToByteArray {
        param (
            [string]$hexString
        )
        if ($null -eq $hexString) {
            return @()
        }

        # remove L"\r\n\t -.,\\/'{}`\"" and whitespace from the hex string
        # this is the same set of characters checked in IsFilteredHex
        $hexString = $hexString -replace "[\r\n\t -.,\\/'{}`"\""]", "" -replace "\s", ""
        if ($hexString.Length -eq 0) {
            return @()
        }

        $byteArray = @()
        for ($i = 0; $i -lt $hexString.Length; $i += 2) {
            try {
                $byteArray += [Convert]::ToByte($hexString.Substring($i, 2), 16)
            } catch {
                Write-Host "Error converting hex string to byte array: $($_.Exception.Message)"
                Write-Host "hexString: $hexString"
                Write-Host "i: $i"
                Write-Host "hexString.Length: $($hexString.Length)"
                # Write the (up to) 8 characters before the error and up to 8 after
                $start = [Math]::Max(0, $i - 8)
                $end = [Math]::Min($hexString.Length, $i + 8)
                Write-Host "hexString.Substring($i, 2): $($hexString.Substring($i, 2))"
                Write-Host "hexString.Substring($start, $end - $start): $($hexString.Substring($start, $end - $start))" 
                break
            }
        }
        return $byteArray
    }
    
    # Iterate over all .dat files in the input directory
    Get-ChildItem -Path $InputDir -Filter *.dat | ForEach-Object {
        $inputFilePath = $_.FullName
        $outputFilePath = Join-Path -Path $OutputDir -ChildPath ($_.BaseName + ".bin")

        # Read the hex data from the input file
        $hexData = Get-Content -Path $inputFilePath -Raw

        Write-Host "Converting $inputFilePath to $outputFilePath"
        # Write-Host "Hex data length: $($hexData.Length)"
        # Write-Host "hexData: $hexData"

        # Convert the hex data to binary data
        $binaryData = Convert-HexStringToByteArray -hexString $hexData
        if ($null -eq $binaryData) {
            $binaryData = @()
        }
        # Write the binary data to the output file
        [System.IO.File]::WriteAllBytes($outputFilePath, $binaryData)
    }
}

# Example usage
$inputDirectory = "$PSScriptRoot\..\UnitTest\SmartViewTestData\In"
$outputDirectory = "$PSScriptRoot\corpus"
Build-FuzzingCorpus -InputDir $inputDirectory -OutputDir $outputDirectory
