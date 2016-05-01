$script:csvData = $null
$script:transformedData = $null
$script:csvGoodData = $null
$script:csvGoodDataLookup = $null

function Get-FirstName {
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [string]
        $fullname
    )

    $index = $fullname.IndexOf(',')

    if ($index -ge 0) {
        Write-Output $fullname.Substring($index + 2)
    } else {
        Write-Output $fullname
    }
}

function Get-LastName {
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [string]
        $fullname
    )

    $index = $fullname.IndexOf(',')

    if ($index -ge 0) {
        Write-Output $fullname.Substring(0, $index)
    } else {
        Write-Output ""
    }
}

Task transform -action {
    $script:transformedData = $script:csvData |
    Select-Object -Property `
        @{ Name = "ID"                  ; Expression = { $_."ID" } },
        @{ Name = "First Name"          ; Expression = { $script:csvGoodDataLookup[$_."Ref Number"]."Full Name" | Get-FirstName } },
        @{ Name = "Last Name"           ; Expression = { $script:csvGoodDataLookup[$_."Ref Number"]."Full Name" | Get-LastName } },
        @{ Name = "Email"               ; Expression = { $script:csvGoodDataLookup[$_."Ref Number"]."E-Mail" } },
        @{ Name = "Phone Number"        ; Expression = { $script:csvGoodDataLookup[$_."Ref Number"]."Phone" } },
        @{ Name = "Ref Number"          ; Expression = { $_."Ref Number" } },
        @{ Name = "Birth Date"          ; Expression = { $_."Birth Date" } },
        @{ Name = "Company Name"        ; Expression = { $script:csvGoodDataLookup[$_."Ref Number"]."Company" } },
        @{ Name = "Address Email"       ; Expression = { $_."Address Email" } },
        @{ Name = "Address Phone Number"; Expression = { $_."Address Phone Number" } },
        @{ Name = "Address ID"          ; Expression = { $_."Address ID" } },
        @{ Name = "Street 1"            ; Expression = { $script:csvGoodDataLookup[$_."Ref Number"]."Street" } },
        @{ Name = "Street 2"            ; Expression = { "" } },
        @{ Name = "City"                ; Expression = { $script:csvGoodDataLookup[$_."Ref Number"]."City" } },
        @{ Name = "State"               ; Expression = { $script:csvGoodDataLookup[$_."Ref Number"]."State" } },
        @{ Name = "Zipcode"             ; Expression = { $script:csvGoodDataLookup[$_."Ref Number"]."ZIP" } },
        @{ Name = "Country"             ; Expression = { "US" } },
        @{ Name = "Address Company Name"; Expression = { $script:csvGoodDataLookup[$_."Ref Number"]."Company" } },
        @{ Name = "Primary Billing"     ; Expression = { $_."Primary Billing" } },
        @{ Name = "Primary Shipping"    ; Expression = { $_."Primary Shipping" } },
        @{ Name = "Customer Group"      ; Expression = { $_."Customer Group" } },
        @{ Name = "Active"              ; Expression = { $_."Active" } },
        @{ Name = "Notes"               ; Expression = { $_."Notes" } },
        @{ Name = "Establishments"      ; Expression = { $_."Establishments" } }
} -depends "load"

Task export -action {
    $script:transformedData | Export-Csv ".\dist\cf-members-export-20160430-valid-ref-num.csv" -Force -NoTypeInformation
} -depends transform

Task load -action {
    $csv = ".\data\cf-revel-members-20160430-valid-ref-num.csv"
    $script:csvData = Import-Csv -Path $csv
    Write-Host "Member Rows: $($script:csvData.Count)"

    $csvGood = ".\data\cf-member-list-20160430.csv"
    $script:csvGoodData = Import-Csv -Path $csvGood | Where-Object { $_."Customer ID" -ne "4.00E+11" }
    Write-Host "Good Data Rows: $($script:csvGoodData.Count)"

    $script:csvGoodDataLookup = ,$script:csvGoodData | New-Dictionary -key "Customer ID"
    Write-Host "Good Data Lookup: $($script:csvGoodDataLookup.Keys.Count)"
    
    $script:csvGoodDataLookup["2099186059"]
}

function New-Dictionary {
    param (
        [string]
        $key,
        [string]
        $value = $null,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        $inList
    )

    $hash = @{}

    $inList | ForEach-Object {
        if ($hash.ContainsKey($_.$key)) {
            Write-Host "The key '$($_.$key)' already exists. Skipping."
        } else {
            if ([string]::IsNullOrEmpty($value)) {
                $hash.Add($_.$key, $_)
            } else {
                $hash.Add($_.$key, $_.$value)
            }
        }
    }

    $hash | Write-Output
}

Task default -depends "export"
