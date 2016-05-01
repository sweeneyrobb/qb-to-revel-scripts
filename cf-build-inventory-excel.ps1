$script:excelFilePath = ".\data\cf-revel-inventory-ceres-20160430.xlsx"
$script:excel = New-Object -Com Excel.Application
$script:workbook = $null
$script:quantityCsv = ".\data\cf-products-20160430.csv"
$script:quantityLookup = $null

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

Task load -action {
  $script:qantityLookup = ,(Import-Csv -Path $quantityCsv |
      Select-Object -Property `
          @{ Name = "Barcode"; Expression = { Get-Barcode -sourceItem $_ }},
          @{ Name = "Quantity"; Expression = { $_."Qty 1" }}) |
      New-Dictionary -key "Barcode" -value "Quantity"
      
  Write-Host "Lookup Count: $($script:qantityLookup.Keys.Count)" -ForegroundColor White
  
  $script:excel.ScreenUpdating = $false
  $script:excel.Interactive = $false
  $script:excel.DisplayAlerts = $false
  
  $script:workbook = $script:excel.Workbooks.Open((Get-Item $script:excelFilePath).FullName)
  $sheet = $script:workbook.Worksheets | Select-Object -First 1

  $usedRange = $sheet.UsedRange

  $usedRange.Rows | Select-Object -Skip 1 | ForEach-Object {
      $barcode = $_.Cells[3].Text
      $_.Cells[14] = $script:qantityLookup[$barcode]
      $_.Cells[17] = "Adjustments from QB POS."
      #Write-Host "Barcode: $barcode, Qty: $($script:qantityLookup[$barcode])" -ForegroundColor Yellow
  }
    
  $script:excel.ScreenUpdating = $true
  $script:excel.Interactive = $true
  $script:excel.DisplayAlerts = $true
  $script:excel.Visible = $true
}

Task default -depends load
