# Define your GitHub token and repository details
$owner = "microsoft"
$repo = "mfcmapi"

# Function to delete a CodeQL analysis
function Delete-CodeQLAnalysisLoop {
    param (
        [string]$analysisUrl
    )

    Write-Host "Request: $analysisUrl"

    # Send the DELETE request to the GitHub API
    $response = gh api --method DELETE -H "Accept: application/vnd.github+json" -H "X-GitHub-Api-Version: 2022-11-28" $analysisUrl

    return $response
}

function Delete-CodeQLAnalysis {
    param (
        [string]$analysisUrl
    )

    # Loop to delete analyses
    while ($analysisUrl) {
        $response = Delete-CodeQLAnalysisLoop -analysisUrl $analysisUrl
        Write-Host "response: $response"

        # Parse the response to get the next analysis URL
        $responseJson = $response | ConvertFrom-Json
        Write-Host "responseJson: $responseJson"
        if ($responseJson.confirm_delete_url) {
            Write-Host "Next response: $responseJson.confirm_delete_url"
            $analysisUrl = $responseJson.confirm_delete_url
        } else {
            $analysisUrl = $null
        }
    }
}

# Fetch the analyses using the GitHub CLI
$analyses = gh api -H "Accept: application/vnd.github+json" -H "X-GitHub-Api-Version: 2022-11-28" /repos/$owner/$repo/code-scanning/analyses --paginate

# Convert the response to JSON
$analysesJson = $analyses | ConvertFrom-Json

# Loop over the analyses and write the ID when deletable is true
foreach ($analysis in $analysesJson) {
    if ($analysis.deletable -eq $true) {
        Write-Output "Deletable analysis url: $($analysis.url)"
        Delete-CodeQLAnalysis -analysisUrl "$($analysis.url)?confirm_delete"
    }
}
