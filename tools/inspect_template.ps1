param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,
    [Parameter(Mandatory = $true)]
    [string]$PreviewPdfPath
)

$ErrorActionPreference = 'Stop'
$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open((Resolve-Path -LiteralPath $WorkbookPath).Path, 0, $true)

    $result = [ordered]@{
        workbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
        sheets = @()
    }

    foreach ($sheet in $workbook.Worksheets) {
        $used = $sheet.UsedRange
        $cells = @()
        for ($row = 1; $row -le $used.Rows.Count; $row++) {
            for ($col = 1; $col -le $used.Columns.Count; $col++) {
                $cell = $used.Cells.Item($row, $col)
                if ($null -ne $cell.Value2 -or $cell.Formula -ne '') {
                    $cells += [ordered]@{
                        address = $cell.Address($false, $false)
                        value = $cell.Value2
                        text = $cell.Text
                        formula = $cell.Formula
                        numberFormat = $cell.NumberFormat
                        interiorColor = $cell.Interior.Color
                        fontColor = $cell.Font.Color
                    }
                }
            }
        }

        $merged = @()
        for ($row = 1; $row -le $used.Rows.Count; $row++) {
            for ($col = 1; $col -le $used.Columns.Count; $col++) {
                $cell = $used.Cells.Item($row, $col)
                if ($cell.MergeCells) {
                    $merged += $cell.MergeArea.Address($false, $false)
                }
            }
        }

        $columnWidths = @()
        for ($col = 1; $col -le $used.Columns.Count; $col++) {
            $columnWidths += [ordered]@{
                column = $used.Columns.Item($col).Address($false, $false)
                width = $used.Columns.Item($col).ColumnWidth
            }
        }

        $rowHeights = @()
        for ($row = 1; $row -le $used.Rows.Count; $row++) {
            $rowHeights += [ordered]@{
                row = $used.Rows.Item($row).Row
                height = $used.Rows.Item($row).RowHeight
            }
        }

        $horizontalPageBreaks = @()
        foreach ($pageBreak in $sheet.HPageBreaks) {
            $horizontalPageBreaks += $pageBreak.Location.Row
        }

        $result.sheets += [ordered]@{
            name = $sheet.Name
            usedRange = $used.Address($false, $false)
            cells = $cells
            mergedRanges = @($merged | Select-Object -Unique)
            columnWidths = $columnWidths
            rowHeights = $rowHeights
            horizontalPageBreakRows = $horizontalPageBreaks
            pageSetup = [ordered]@{
                orientation = $sheet.PageSetup.Orientation
                paperSize = $sheet.PageSetup.PaperSize
                zoom = $sheet.PageSetup.Zoom
                fitToPagesWide = $sheet.PageSetup.FitToPagesWide
                fitToPagesTall = $sheet.PageSetup.FitToPagesTall
                printArea = $sheet.PageSetup.PrintArea
                printTitleRows = $sheet.PageSetup.PrintTitleRows
                centerHorizontally = $sheet.PageSetup.CenterHorizontally
                centerVertically = $sheet.PageSetup.CenterVertically
            }
        }
    }

    $previewPdfFullPath = [IO.Path]::GetFullPath((Join-Path (Get-Location) $PreviewPdfPath))
    $previewDirectory = Split-Path -Parent $previewPdfFullPath
    if ($previewDirectory) {
        New-Item -ItemType Directory -Force -Path $previewDirectory | Out-Null
    }
    $workbook.ExportAsFixedFormat(0, $previewPdfFullPath)
    $result | ConvertTo-Json -Depth 8
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    if ($null -ne $excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
