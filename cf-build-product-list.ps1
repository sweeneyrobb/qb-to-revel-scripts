$script:csvData = $null
$script:transformedData = $null
$script:vendorLookup = $null
$script:categoryLookup = $null
$script:subCategoryLookup = $null

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

    $newList = New-Object System.Collections.Generic.List[System.Object]

    $productList |
    Group-Object -Property 'Item Name' |
    ForEach-Object {
        if ($_.Count -gt 1) {
            Write-Host "Creating Parent for group $($_.Name), count: $($_.Count)..." -ForegroundColor Yellow
            $source = $_.Group[0]
            $parent = New-ProductParent -sourceItem $source
            $newList.Add($parent)

            $_.Group | ForEach-Object {
                $name = $_."Item Name"
                $_."Item Name" += " $($_.Attribute) $($_.Size)"
                $newList.Add((Add-AttributeTypeAndParent -parentId $parent."Item Number" -attributeType "Child" -row $_))
            }
        } else {
            $newList.Add((Add-AttributeTypeAndParent -parentId "" -attributeType "None" -row $_.Group[0]))
        }
    }

    $newList
}

function Get-CategoryBasedName {
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        $product
    )
    $name = $product."Item Name"
    $deptName = $product."Department Name"
    $category = $script:categoryLookup[$deptName]

    Write-Output "$category $name"
}


function Remove-PaddedZeroes {
    param (
        [string]
        $inString
    )

    if (-not [string]::IsNullOrEmpty($inString) -and $inString.Length -gt 0 -and $inString.StartsWith('0')) {
        Remove-PaddedZeroes -inString $inString.Substring(1)
    } else {
        $inString
    }
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
                $_."Item Name" = Get-CategoryBasedName -product $_
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

    ,$newList
}

function New-ProductParent {
    param(
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
    Add-Member -Name "Attribute parent"        -Value ""                                    -MemberType NoteProperty -PassThru |
    Add-Member -Name "Attribute type"          -Value "Parent"                              -MemberType NoteProperty -PassThru |
    Write-Output
}

Task transform -action {
    $script:transformedData = $script:csvData |
    Where-Object { $script:categoryLookup[$_."Department Name"] -ne "Not Used" } |
    Where-Object { $_."Item Name" -ne "Accessories" -and
                   $_."Item Name" -ne "Accessories 10" -and
                   $_."Item Name" -ne "Accessories 20" -and
                   $_."Item Name" -ne "Accessories 30" -and
                   $_."Item Name" -ne "Accessories 40" -and
                   $_."Item Name" -ne "Accessories 50" -and
                   $_."Item Name" -ne "Accessories 60" }

    $script:transformedData = , $script:transformedData | Process-ProductGroupForDuplicateCategory | Process-ProductGroup

    Write-Host "Transformed Rows: $($script:transformedData.Count)"
} -depends load

function Get-Barcode {
    param(
        $sourceItem
    )

    if ($sourceItem.UPC -ne $null -and $sourceItem.UPC -ne "") {
        Remove-PaddedZeroes -inString $sourceItem.UPC
    } else {
        "{0:000000}" -f [int]$sourceItem."Item Number"
    }
}

Task transform-select -action {
    $script:transformedData = $script:transformedData | Select-Object -Property `
        @{ Name = "Product Class"; Expression = { "Apparel" } },
        @{ Name = "Product Category"; Expression = { $script:categoryLookup[$_."Department Name"] } },
        @{ Name = "Product Subcategory"; Expression = {  $script:subCategoryLookup[$_."Department Name"] } },
        @{ Name = "Product Name"; Expression = { $_."Item Name" } },
        @{ Name = "Product Description"; Expression = { $_."Item Description" } },
        @{ Name = "Price"; Expression = { $_."Regular Price" } },
        @{ Name = "Cost"; Expression = { $_."Order Cost" } },
        @{ Name = "SKU"; Expression = { "" } },
        @{ Name = "Barcode"; Expression = { Get-Barcode -sourceItem $_ } },
        @{ Name = "Active"; Expression = { "Yes" } },
        @{ Name = "Attribute type"; Expression = { $_."Attribute type" } },
        @{ Name = "Attribute parent"; Expression = { $_."Attribute parent" } },
        @{ Name = "Attribute 1 name"; Expression = { "" } },
        @{ Name = "Attribute 2 name"; Expression = { "" } },
        @{ Name = "Attribute 1"; Expression = { "Size" } },
        @{ Name = "Attribute 2"; Expression = { "Style" } },
        @{ Name = "Attribute value 1"; Expression = { $_.Size } },
        @{ Name = "Attribute value 2"; Expression = { $_.Attribute } },
        @{ Name = "Default inventory"; Expression = { "No" } },
        @{ Name = "BIN number"; Expression = { "" } },
        @{ Name = "Unit of Measurement"; Expression = { "Unit" } },
        @{ Name = "Primary stock unit name"; Expression = { "Unit" } },
        @{ Name = "Conversion factor"; Expression = { "1" } },
        @{ Name = "Track in inventory"; Expression = { "Yes" } },
        @{ Name = "Inventory Threshold"; Expression = { "" } },
        @{ Name = "Primary vendor ID"; Expression = { $script:vendorLookup[$_."Vendor Name"] } },
        @{ Name = "Id"; Expression = { $_."Item Number" } },
        @{ Name = "Point Value"; Expression = { [int] $_."Regular Price" } }
} -depends "transform"

Task export -action {
    $script:transformedData | Export-Csv ".\dist\test.csv" -Force -NoTypeInformation
} -depends transform-select

Task load -action {
    $csv = ".\data\cf-products-20160428.csv"
    $vendorCsv = ".\data\cf-vendors.csv"
    $categoriesCsv = ".\data\cf-categories.csv"

    $script:csvData = Import-Csv -Path $csv
    Write-Host "Product Rows: $($script:csvData.Count)"

    $script:vendorLookup = ,(Import-Csv -Path $vendorCsv) |
                           New-Dictionary -key "Vendor name" -value "ID"

    Write-Host "Vendor Rows: $($script:vendorLookup.Keys.Count)"

    $cdata = ,(Import-Csv -Path $categoriesCsv)
    $script:categoryLookup = $cdata |
                             New-Dictionary -key "Department" -value "Category"

    Write-Host "Category Rows: $($script:categoryLookup.Keys.Count)"

    $script:subCategoryLookup = $cdata |
                                New-Dictionary -key "Department" -value "Sub Category"

    Write-Host "Sub Category Rows: $($script:subCategoryLookup.Keys.Count)"
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
