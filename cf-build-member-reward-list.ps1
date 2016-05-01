$script:csvData = $null
$script:transformedData = $null
$script:memberIdLookup = $null

function New-Dictionary {
    param (
        [string]
        $key,
        [string]
        $value,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        $inList
    )

    $hash = @{}

    $inList | ForEach-Object {
        if (-not $hash.ContainsKey($_.$key)) {
            $hash.Add($_.$key, $_.$value)
        } else {
            Write-Host "Key '$($_.$key)' Already Exists. Skipping." -ForegroundColor Yellow
        }
    }

    $hash | Write-Output
}

Task export -action {
    $script:transformedData | Export-Csv ".\dist\cf-member-rewards-export-20160430.csv" -Force -NoTypeInformation
} -depends transform

function Get-RewardPoints {
    param ($row)
    
    $points = (([int]$row."# Rewards Available" * 200)) + (200 - ([int]$row."Purchase $ to Next Reward"))
    
    $points | Write-Output 
}

Task transform -action {
    $script:transformedData = $script:csvData | Select-Object -Property `
      @{ Name = "Number";               Expression = { $_."Customer ID" } },
      @{ Name = "Points By Visits";     Expression = { "0" } },
      @{ Name = "Points By Purchases";  Expression = { Get-RewardPoints -row $_ } },
      @{ Name = "Current Points";       Expression = { Get-RewardPoints -row $_ } }
} -depends load

Task load -action {
    $csv = ".\data\cf-member-list-20160430.csv"
    $script:csvData = Import-Csv -Path $csv | Where-Object { $_."Customer ID" -ne "4.00E+11" }
    Write-Host "Member Reward Rows: $($script:csvData.Count)"

    <#
    $csvExport = ".\data\cf-revel-export-20160429.csv"
    $script:memberIdLookup = ,(Import-Csv -Path $csvExport) |
                               New-Dictionary -key "Ref Number" -value "ID"

    Write-Host "Member Lookup Rows: $($script:memberIdLookup.Keys.Count)"
    #>
}

Task default -depends export
