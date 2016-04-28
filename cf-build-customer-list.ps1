$script:csvData = $null
$script:transformedData = $null

function Add-AttributeTypeAndParent {
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        $row,
        $parentId,
        $attributeType
    )

    $row |
        Add-Member -Name "Attribute parent" -Value $parentId -MemberType NoteProperty -PassThru |
        Add-Member -Name "Attribute type" -Value $attributeType -MemberType NoteProperty -PassThru |
        Write-Output
}

function Process-ProductGroup {
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [Array]
        $productList
    )

    $newList = @()

    $productList |
    Group-Object -Property "Item Name" |
    ForEach-Object {
        if ($_.Count -gt 1) {
            $parent = $_.Group[0] | New-ProductParent
            $newList += $parent

            $_.Group | ForEach-Object {
                $_."Item Name" += " $($_.Attribute) $($_.Size)"
                $newList += $_ | Add-AttributeTypeAndParent -parentId $parent."Item Number" -attributeType "Child"
            }
        } else {
            $newList += $_.Group[0] | Add-AttributeTypeAndParent -parentId "" -attributeType "None"
        }
    }

    , $newList
}

function Get-CategoryBasedName {
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        $product
    )

    Write-Output $script:categoryLookup[$product."Department Name"] $product."Item Name"
}

function Process-ProductGroupForDuplicateCategory {
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        $products
    )
    $newList = @()

    $products |
    Group-Object -Property "Item Name" |
    ForEach-Object {
        $groupedByDept = $_.Group | Group-Object { $script:categoryLookup[$_."Department Name"] }

        <#
        if ($groupedByDept.GetType().IsArray) {
            Write-Host "'$($_.Name)' count: $($groupedByDept.Count) isArray: $($groupedByDept.GetType().IsArray)" -ForegroundColor Red
            $groupedByDept | ForEach-Object {
                Write-Host "Categories: $($_.Name)" -ForegroundColor DarkRed
            }
        }
        #>

        if ($_.Count -gt 1 -and $groupedByDept.GetType().IsArray -and $groupedByDept.Count -gt 1 ) {
            $_.Group | ForEach-Object {
                $_."Item Name" =  $_ | Get-CategoryBasedName

                <#
                $newName = $_."Item Name" + "old dept: " + $_."Department Name"
                Write-Host " - New Name: $newName" -ForegroundColor Red
                #>

                $newList += $_
            }
        } else {
            $newList += $_.Group
        }
    }

    , $newList
}

function New-ProductParent {
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        $sourceItem
    )

    $parentItemNumber = ([int]$sourceItem."Item Number") + 100000

    New-Object PSObject |
    Add-Member -Name "Item Number"             -value $parentItemNumber                     -MemberType NoteProperty -PassThru |
    Add-Member -Name "Item Name"               -value $sourceItem."Item Name"               -MemberType NoteProperty -PassThru |
    Add-Member -Name "Item Description"        -value $sourceItem."Item Description"        -MemberType NoteProperty -PassThru |
    Add-Member -Name "Alternate Lookup"        -value $sourceItem."Alternate Lookup"        -MemberType NoteProperty -PassThru |
    Add-Member -Name "Attribute"               -value ""                                    -MemberType NoteProperty -PassThru |
    Add-Member -Name "Size"                    -value ""                                    -MemberType NoteProperty -PassThru |
    Add-Member -Name "Average Unit Cost"       -value $sourceItem."Average Unit Cost"       -MemberType NoteProperty -PassThru |
    Add-Member -Name "Regular Price"           -value $sourceItem."Regular Price"           -MemberType NoteProperty -PassThru |
    Add-Member -Name "MSRP"                    -value $sourceItem."MSRP"                    -MemberType NoteProperty -PassThru |
    Add-Member -Name "Custom Price 1"          -value $sourceItem."Custom Price 1"          -MemberType NoteProperty -PassThru |
    Add-Member -Name "Custom Price 2"          -value $sourceItem."Custom Price 2"          -MemberType NoteProperty -PassThru |
    Add-Member -Name "Custom Price 3"          -value $sourceItem."Custom Price 3"          -MemberType NoteProperty -PassThru |
    Add-Member -Name "Custom Price 4"          -value $sourceItem."Custom Price 4"          -MemberType NoteProperty -PassThru |
    Add-Member -Name "Tax Code"                -value $sourceItem."Tax Code"                -MemberType NoteProperty -PassThru |
    Add-Member -Name "UPC"                     -value ""                                    -MemberType NoteProperty -PassThru |
    Add-Member -Name "Order Cost"              -value $sourceItem."Order Cost"              -MemberType NoteProperty -PassThru |
    Add-Member -Name "Item Type"               -value $sourceItem."Item Type"               -MemberType NoteProperty -PassThru |
    Add-Member -Name "Base Unit of Measure"    -value $sourceItem."Base Unit of Measure"    -MemberType NoteProperty -PassThru |
    Add-Member -Name "Company Reorder Point"   -value $sourceItem."Company Reorder Point"   -MemberType NoteProperty -PassThru |
    Add-Member -Name "Print Tags"              -value $sourceItem."Print Tags"              -MemberType NoteProperty -PassThru |
    Add-Member -Name "Unorderable"             -value $sourceItem."Unorderable"             -MemberType NoteProperty -PassThru |
    Add-Member -Name "Serial Tracking"         -value $sourceItem."Serial Tracking"         -MemberType NoteProperty -PassThru |
    Add-Member -Name "Eligible for Commission" -value $sourceItem."Eligible for Commission" -MemberType NoteProperty -PassThru |
    Add-Member -Name "Department Name"         -value $sourceItem."Department Name"         -MemberType NoteProperty -PassThru |
    Add-Member -Name "Category"                -value $sourceItem."Category"                -MemberType NoteProperty -PassThru |
    Add-Member -Name "Sub Category"            -value $sourceItem."Sub Category"            -MemberType NoteProperty -PassThru |
    Add-Member -Name "Department Code"         -value $sourceItem."Department Code"         -MemberType NoteProperty -PassThru |
    Add-Member -Name "Vendor Name"             -value $sourceItem."Vendor Name"             -MemberType NoteProperty -PassThru |
    Add-Member -Name "Vendor Code"             -value $sourceItem."Vendor Code"             -MemberType NoteProperty -PassThru |
    Add-Member -Name "Manufacturer"            -value $sourceItem."Manufacturer"            -MemberType NoteProperty -PassThru |
    Add-Member -Name "Qty 1"                   -value 0                                     -MemberType NoteProperty -PassThru |
    Add-Member -Name "Qty 2"                   -value 0                                     -MemberType NoteProperty -PassThru |
    Add-Member -Name "Qty 3"                   -value 0                                     -MemberType NoteProperty -PassThru |
    Add-Member -Name "Eligible for Rewards"    -value $sourceItem."Eligible for Rewards"    -MemberType NoteProperty -PassThru |
    Add-Member -Name "Vendor Id"               -value $sourceItem."Vendor Id"               -MemberType NoteProperty -PassThru |
    Add-Member -Name "Attribute parent"        -Value $parentItemNumber                     -MemberType NoteProperty -PassThru |
    Add-Member -Name "Attribute type"          -Value "Parent"                              -MemberType NoteProperty -PassThru |
    Write-Output
}

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

Task transform-select -action {
    $script:transformedData = $script:csvData |
    Select-Object -Property `
        @{ Name = "ID"                   ; Expression = { "" } },
        @{ Name = "First Name"           ; Expression = { $_."First Name" } },
        @{ Name = "Last Name"            ; Expression = { $_."Last Name" } },
        @{ Name = "Email"                ; Expression = { $_."EMail" } },
        @{ Name = "Phone Number"         ; Expression = { $_."Phone 1" } },
        @{ Name = "Ref Number"           ; Expression = { $_."Phone 1" } },
        @{ Name = "Birth Date"           ; Expression = { "" } },
        @{ Name = "Company Name"         ; Expression = { $_."Company" } },
        @{ Name = "Address Email"        ; Expression = { "" } },
        @{ Name = "Address Phone Number" ; Expression = { "" } },
        @{ Name = "Address ID"           ; Expression = { $_."" } },
        @{ Name = "Street 1"             ; Expression = { $_."Street" } },
        @{ Name = "Street 2"             ; Expression = { $_."Street2" } },
        @{ Name = "City"                 ; Expression = { $_."City" } },
        @{ Name = "State"                ; Expression = { $_."State" } },
        @{ Name = "Zipcode"              ; Expression = { $_."ZIP" } },
        @{ Name = "Country"              ; Expression = { "US" } },
        @{ Name = "Address Company Name" ; Expression = { "" } },
        @{ Name = "Primary Billing"      ; Expression = { "Yes" } },
        @{ Name = "Primary Shipping"     ; Expression = { "Yes" } },
        @{ Name = "Customer Group"       ; Expression = { "" } },
        @{ Name = "Active"               ; Expression = { "Yes" } },
        @{ Name = "Notes"                ; Expression = { "" } },
        @{ Name = "Establishments"       ; Expression = { "3,4" } }
} -depends "load"

Task export -action {
    $script:transformedData | Export-Csv ".\dist\cf-customers-export-non-rewards-20160428.csv" -Force -NoTypeInformation
} -depends transform-select

Task load -action {
    $csv = ".\data\cf-customers-non-rewards-20160428.csv"

    $script:csvData = Import-Csv -Path $csv
    Write-Host "Member Rows: $($script:csvData.Count)"
}

Task find-departmentitemconflicts -action {
    $script:csvData | Group-Object -Property "Item Name" | Where-Object { $_.Count -gt 1 } | ForEach-Object {
        $_ | Where-Object { ($_.Group | Select-Object -Property "Department Name" -Unique).Count -gt 1 } | ForEach-Object {
            $_.Group | Select-Object -Property "Item Name", "Department Name"
        }
    } | fl
} -depends load

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
        $hash.Add($_.$key, $_.$value)
    }

    $hash | Write-Output
}

Task default -depends "export"
