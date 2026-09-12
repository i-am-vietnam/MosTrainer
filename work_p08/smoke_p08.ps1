$ErrorActionPreference = 'Stop'

$repo = 'C:\Users\HUYNH HAU\source\repos'
$starter = Join-Path $repo 'MosTrainer\Projects\Excel2019_P08\starter.xlsx'
$core = Join-Path $repo 'MosTrainer.Core\bin\Debug\MosTrainer.Core.dll'
$excelAssembly = Join-Path $repo 'MosTrainer.Excel\bin\Debug\MosTrainer.Excel.dll'
[void][Reflection.Assembly]::LoadFrom($core)
[void][Reflection.Assembly]::LoadFrom($excelAssembly)

function Release-Com($value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)
    }
}

function New-Case([string]$name, [scriptblock]$mutate) {
    $path = Join-Path $env:TEMP ("MosTrainer-P08-smoke-{0}.xlsx" -f $name)
    Copy-Item -LiteralPath $starter -Destination $path -Force
    $app = $null; $wb = $null
    try {
        $app = New-Object -ComObject Excel.Application
        $app.Visible = $false
        $app.DisplayAlerts = $false
        $wb = $app.Workbooks.Open($path)
        $null = & $mutate $wb
        $wb.Save()
    }
    finally {
        if ($null -ne $wb) { $wb.Close($false) }
        if ($null -ne $app) { $app.Quit() }
        Release-Com $wb; Release-Com $app
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    return $path
}

function Grade([string]$path, [string]$method, [object[]]$arguments) {
    $controller = New-Object MosTrainer.Excel.ExcelController
    try {
        $controller.OpenWorkbook($path)
        return [bool]$controller.GetType().GetMethod($method).Invoke($controller, $arguments)
    }
    finally {
        $controller.Dispose()
    }
}

$results = [System.Collections.Generic.List[object]]::new()
function Check([string]$task, [string]$case, [bool]$expected, [string]$path, [string]$method, [object[]]$arguments) {
    $actual = Grade $path $method $arguments
    $results.Add([pscustomobject]@{Task=$task; Case=$case; Expected=$expected; Actual=$actual; OK=($actual -eq $expected)})
}

$unchanged = New-Case 'unchanged' { param($wb) }

$p = New-Case 't01-pass' { param($wb)
    $ws=$wb.Worksheets.Item('Forecast'); $t=$ws.ListObjects.Item('Table5');
    for($r=1;$r -le $t.DataBodyRange.Rows.Count;$r++){ $row=$t.DataBodyRange.Row+$r-1; $ws.Cells.Item($row,3).Formula = "=B$row*Q2_Increase" }
    Release-Com $t; Release-Com $ws
}
Check 'T01' 'correct direct formulas' $true $p 'TableColumnFormulaMultipliesNamedRange' @('Forecast','Table5','Quarter 2','Quarter 1','Q2_Increase','B16')
Check 'T01' 'untouched' $false $unchanged 'TableColumnFormulaMultipliesNamedRange' @('Forecast','Table5','Quarter 2','Quarter 1','Q2_Increase','B16')

$p = New-Case 't02-pass' { param($wb)
    $ws=$wb.Worksheets.Item('Suppliers'); $t=$ws.ListObjects.Item('Table2'); $t.ListRows.Item(2).Delete(); Release-Com $t; Release-Com $ws
}
$ids=[string[]]@('P020','P029','P027','P043','P014','P016','P017','P034','P004')
Check 'T02' 'delete table row' $true $p 'TableRowContainingTextDeletedPreserveOutside' @('Suppliers','Table2','Wooden Toys',9,'A1:G10','A1:L11',$ids)
Check 'T02' 'untouched' $false $unchanged 'TableRowContainingTextDeletedPreserveOutside' @('Suppliers','Table2','Wooden Toys',9,'A1:G10','A1:L11',$ids)

$p = New-Case 't03-pass' { param($wb)
    $ws=$wb.Worksheets.Item('First half of the year'); $r=$ws.Range('A4:A11'); $r.HorizontalAlignment=-4131; $r.IndentLevel=1; Release-Com $r; Release-Com $ws
}
Check 'T03' 'left indent one' $true $p 'RangeAlignmentIndentEquals' @('First half of the year','A4:A11','Left',1)
Check 'T03' 'untouched' $false $unchanged 'RangeAlignmentIndentEquals' @('First half of the year','A4:A11','Left',1)

$p = New-Case 't04-pass' { param($wb)
    $ws=$wb.Worksheets.Item('First half of the year'); $r=$ws.Range('J4:J11'); [void]$r.SparklineGroups.Add(3,"'First half of the year'!B4:G11"); Release-Com $r; Release-Com $ws
}
Check 'T04' 'win loss exact rows' $true $p 'SparklinesByRangeAndType' @('First half of the year','J4:J11','B4:G11','WinLoss')
Check 'T04' 'untouched' $false $unchanged 'SparklinesByRangeAndType' @('First half of the year','J4:J11','B4:G11','WinLoss')

$p = New-Case 't05-pass' { param($wb)
    $ws=$wb.Worksheets.Item('First half of the year'); $t=$ws.ListObjects.Item('Q1_Sales'); $t.ShowTotals=$true
    foreach($h in @('January','February','March','April','May','June','Total')){$t.ListColumns.Item($h).TotalsCalculation=1}
    Release-Com $t; Release-Com $ws
}
$headers=[string[]]@('January','February','March','April','May','June','Total')
Check 'T05' 'totals row sums' $true $p 'TableTotalRowSumsByHeaders' @('First half of the year','Q1_Sales',$headers)
Check 'T05' 'untouched' $false $unchanged 'TableTotalRowSumsByHeaders' @('First half of the year','Q1_Sales',$headers)

$p = New-Case 't06-pass' { param($wb)
    $ws=$wb.Worksheets.Item('First half of the year'); for($row=4;$row -le 11;$row++){$ws.Cells.Item($row,9).Formula=('=COUNTBLANK(B{0}:G{0})' -f $row)}; Release-Com $ws
}
$months=[string[]]@('January','February','March','April','May','June')
Check 'T06' 'countblank direct formulas' $true $p 'CountBlankFormulaByHeaders' @('First half of the year','Q1_Sales','Inactive',$months)
Check 'T06' 'untouched' $false $unchanged 'CountBlankFormulaByHeaders' @('First half of the year','Q1_Sales','Inactive',$months)

$p = New-Case 't07-pass' { param($wb)
    $ws=$wb.Worksheets.Item('Top Toys Category'); $t=$ws.ListObjects.Item('Table4'); $s=$t.Sort; $s.SortFields.Clear();
    [void]$s.SortFields.Add($t.ListColumns.Item('Toys').DataBodyRange,0,1); [void]$s.SortFields.Add($t.ListColumns.Item('Total Sales').DataBodyRange,0,2); $s.Header=1; $s.Apply()
    Release-Com $s; Release-Com $t; Release-Com $ws
}
$sortHeaders=[string[]]@('Toys','Total Sales'); $sortOrders=[string[]]@('Ascending','Descending')
Check 'T07' 'two persisted sort fields' $true $p 'TableMultiLevelSortStateEquals' @('Top Toys Category','Table4','A2:B12',$sortHeaders,$sortOrders)
Check 'T07' 'untouched' $false $unchanged 'TableMultiLevelSortStateEquals' @('Top Toys Category','Table4','A2:B12',$sortHeaders,$sortOrders)

$p = New-Case 't08-pass' { param($wb)
    $ws=$wb.Worksheets.Item('Last half of the year'); $co=$ws.ChartObjects().Item('Chart 1'); $co.Chart.ApplyLayout(3); Release-Com $co; Release-Com $ws
}
Check 'T08' 'quick layout 3' $true $p 'ChartQuickLayoutEquals' @('Last half of the year','Chart 1',3,51,'A4:A11','B3:H3','B4:H11')
Check 'T08' 'untouched' $false $unchanged 'ChartQuickLayoutEquals' @('Last half of the year','Chart 1',3,51,'A4:A11','B3:H3','B4:H11')

# Additional false-positive and equivalent-form tests.
$p = New-Case 't01-structured' { param($wb)
    $ws=$wb.Worksheets.Item('Forecast'); $t=$ws.ListObjects.Item('Table5'); $target=$t.ListColumns.Item(3).DataBodyRange; $sourceName=[string]$t.ListColumns.Item(2).Name
    $target.Formula=('=[@[{0}]]*Q2_Increase' -f $sourceName); Release-Com $target; Release-Com $t; Release-Com $ws
}
Check 'T01' 'structured formula equivalent' $true $p 'TableColumnFormulaMultipliesNamedRange' @('Forecast','Table5','Quarter 2','Quarter 1','Q2_Increase','B16')
$p = New-Case 't01-cellref' { param($wb)
    $ws=$wb.Worksheets.Item('Forecast'); for($row=4;$row -le 11;$row++){$ws.Cells.Item($row,3).Formula=('=B{0}*$B$16' -f $row)}; Release-Com $ws
}
Check 'T01' 'named range replaced by cell ref' $false $p 'TableColumnFormulaMultipliesNamedRange' @('Forecast','Table5','Quarter 2','Quarter 1','Q2_Increase','B16')
$p = New-Case 't01-literal' { param($wb)
    $ws=$wb.Worksheets.Item('Forecast'); for($row=4;$row -le 11;$row++){$ws.Cells.Item($row,3).Formula=('=B{0}*1.03' -f $row)}; Release-Com $ws
}
Check 'T01' 'named range replaced by literal' $false $p 'TableColumnFormulaMultipliesNamedRange' @('Forecast','Table5','Quarter 2','Quarter 1','Q2_Increase','B16')
$p = New-Case 't01-typed' { param($wb)
    $ws=$wb.Worksheets.Item('Forecast'); for($row=4;$row -le 11;$row++){$ws.Cells.Item($row,3).Value2=[double]$ws.Cells.Item($row,2).Value2*1.03}; Release-Com $ws
}
Check 'T01' 'typed results' $false $p 'TableColumnFormulaMultipliesNamedRange' @('Forecast','Table5','Quarter 2','Quarter 1','Q2_Increase','B16')
$p = New-Case 't01-wrongrow' { param($wb) $ws=$wb.Worksheets.Item('Forecast'); for($row=4;$row-le11;$row++){$ws.Cells.Item($row,3).Formula=('=B{0}*Q2_Increase'-f$row)}; $ws.Cells.Item(11,3).Formula='=B10*Q2_Increase'; Release-Com $ws }
Check 'T01' 'one row references wrong source row' $false $p 'TableColumnFormulaMultipliesNamedRange' @('Forecast','Table5','Quarter 2','Quarter 1','Q2_Increase','B16')
$p = New-Case 't01-wrongcolumn' { param($wb) $ws=$wb.Worksheets.Item('Forecast'); for($row=4;$row-le11;$row++){$ws.Cells.Item($row,3).Formula=('=A{0}*Q2_Increase'-f$row)}; Release-Com $ws }
Check 'T01' 'wrong source column' $false $p 'TableColumnFormulaMultipliesNamedRange' @('Forecast','Table5','Quarter 2','Quarter 1','Q2_Increase','B16')

$p = New-Case 't02-entirerow' { param($wb) $ws=$wb.Worksheets.Item('Suppliers'); $ws.Rows.Item(3).Delete(); Release-Com $ws }
Check 'T02' 'delete whole worksheet row' $false $p 'TableRowContainingTextDeletedPreserveOutside' @('Suppliers','Table2','Wooden Toys',9,'A1:G10','A1:L11',$ids)
$p = New-Case 't02-clear' { param($wb) $ws=$wb.Worksheets.Item('Suppliers'); $ws.Cells.Item(3,3).ClearContents(); Release-Com $ws }
Check 'T02' 'clear matching text only' $false $p 'TableRowContainingTextDeletedPreserveOutside' @('Suppliers','Table2','Wooden Toys',9,'A1:G10','A1:L11',$ids)
$p = New-Case 't02-wrongrow' { param($wb) $ws=$wb.Worksheets.Item('Suppliers'); $t=$ws.ListObjects.Item('Table2'); $t.ListRows.Item(1).Delete(); Release-Com $t; Release-Com $ws }
Check 'T02' 'delete another table row' $false $p 'TableRowContainingTextDeletedPreserveOutside' @('Suppliers','Table2','Wooden Toys',9,'A1:G10','A1:L11',$ids)

$p = New-Case 't03-zero' { param($wb) $ws=$wb.Worksheets.Item('First half of the year'); $r=$ws.Range('A4:A11'); $r.HorizontalAlignment=-4131; $r.IndentLevel=0; Release-Com $r; Release-Com $ws }
Check 'T03' 'left but indent zero' $false $p 'RangeAlignmentIndentEquals' @('First half of the year','A4:A11','Left',1)
$p = New-Case 't03-missing' { param($wb) $ws=$wb.Worksheets.Item('First half of the year'); $r=$ws.Range('A4:A10'); $r.HorizontalAlignment=-4131; $r.IndentLevel=1; Release-Com $r; Release-Com $ws }
Check 'T03' 'one cell missing' $false $p 'RangeAlignmentIndentEquals' @('First half of the year','A4:A11','Left',1)
$p = New-Case 't03-two' { param($wb) $ws=$wb.Worksheets.Item('First half of the year'); $r=$ws.Range('A4:A11'); $r.HorizontalAlignment=-4131; $r.IndentLevel=2; Release-Com $r; Release-Com $ws }
Check 'T03' 'indent two' $false $p 'RangeAlignmentIndentEquals' @('First half of the year','A4:A11','Left',1)

$p = New-Case 't04-column' { param($wb) $ws=$wb.Worksheets.Item('First half of the year'); $r=$ws.Range('J4:J11'); [void]$r.SparklineGroups.Add(2,"'First half of the year'!B4:G11"); Release-Com $r; Release-Com $ws }
Check 'T04' 'column instead of win loss' $false $p 'SparklinesByRangeAndType' @('First half of the year','J4:J11','B4:G11','WinLoss')
$p = New-Case 't04-wrongdata' { param($wb) $ws=$wb.Worksheets.Item('First half of the year'); $r=$ws.Range('J4:J11'); [void]$r.SparklineGroups.Add(3,"'First half of the year'!C4:H11"); Release-Com $r; Release-Com $ws }
Check 'T04' 'wrong source range' $false $p 'SparklinesByRangeAndType' @('First half of the year','J4:J11','B4:G11','WinLoss')
$p = New-Case 't04-missing' { param($wb) $ws=$wb.Worksheets.Item('First half of the year'); $r=$ws.Range('J4:J10'); [void]$r.SparklineGroups.Add(3,"'First half of the year'!B4:G10"); Release-Com $r; Release-Com $ws }
Check 'T04' 'one sparkline missing' $false $p 'SparklinesByRangeAndType' @('First half of the year','J4:J11','B4:G11','WinLoss')
$p = New-Case 't04-line' { param($wb) $ws=$wb.Worksheets.Item('First half of the year'); $r=$ws.Range('J4:J11'); [void]$r.SparklineGroups.Add(1,"'First half of the year'!B4:G11"); Release-Com $r; Release-Com $ws }
Check 'T04' 'line instead of win loss' $false $p 'SparklinesByRangeAndType' @('First half of the year','J4:J11','B4:G11','WinLoss')
$p = New-Case 't04-wronglocation' { param($wb) $ws=$wb.Worksheets.Item('First half of the year'); $r=$ws.Range('I4:I11'); [void]$r.SparklineGroups.Add(3,"'First half of the year'!B4:G11"); Release-Com $r; Release-Com $ws }
Check 'T04' 'wrong location range' $false $p 'SparklinesByRangeAndType' @('First half of the year','J4:J11','B4:G11','WinLoss')

$p = New-Case 't05-onlyshow' { param($wb) $ws=$wb.Worksheets.Item('First half of the year'); $t=$ws.ListObjects.Item('Q1_Sales'); $t.ShowTotals=$true; Release-Com $t; Release-Com $ws }
Check 'T05' 'totals row without all requested sums' $false $p 'TableTotalRowSumsByHeaders' @('First half of the year','Q1_Sales',$headers)
$p = New-Case 't05-onewrong' { param($wb)
    $ws=$wb.Worksheets.Item('First half of the year'); $t=$ws.ListObjects.Item('Q1_Sales'); $t.ShowTotals=$true
    foreach($h in @('January','February','March','April','May','June','Total')){$t.ListColumns.Item($h).TotalsCalculation=1}; $t.ListColumns.Item('June').TotalsCalculation=2
    Release-Com $t; Release-Com $ws
}
Check 'T05' 'one totals calculation wrong' $false $p 'TableTotalRowSumsByHeaders' @('First half of the year','Q1_Sales',$headers)
$p = New-Case 't05-aprilwrong' { param($wb) $ws=$wb.Worksheets.Item('First half of the year'); $t=$ws.ListObjects.Item('Q1_Sales'); $t.ShowTotals=$true; foreach($h in @('January','February','March','April','May','June','Total')){$t.ListColumns.Item($h).TotalsCalculation=1}; $t.ListColumns.Item('April').TotalsCalculation=0; Release-Com $t; Release-Com $ws }
Check 'T05' 'April not Sum' $false $p 'TableTotalRowSumsByHeaders' @('First half of the year','Q1_Sales',$headers)
$p = New-Case 't05-totalwrong' { param($wb) $ws=$wb.Worksheets.Item('First half of the year'); $t=$ws.ListObjects.Item('Q1_Sales'); $t.ShowTotals=$true; foreach($h in @('January','February','March','April','May','June','Total')){$t.ListColumns.Item($h).TotalsCalculation=1}; $t.ListColumns.Item('Total').TotalsCalculation=0; Release-Com $t; Release-Com $ws }
Check 'T05' 'six month Total not Sum' $false $p 'TableTotalRowSumsByHeaders' @('First half of the year','Q1_Sales',$headers)
$p = New-Case 't05-typed' { param($wb) $ws=$wb.Worksheets.Item('First half of the year'); $t=$ws.ListObjects.Item('Q1_Sales'); $t.ShowTotals=$true; $vals=@(48,23,43,20,43,40,217); $cols=@('January','February','March','April','May','June','Total'); for($i=0;$i-lt$cols.Count;$i++){$index=$t.ListColumns.Item($cols[$i]).Index;$t.TotalsRowRange.Cells.Item(1,$index).Value2=$vals[$i]}; Release-Com $t; Release-Com $ws }
Check 'T05' 'typed totals instead of calculation' $false $p 'TableTotalRowSumsByHeaders' @('First half of the year','Q1_Sales',$headers)

$p = New-Case 't06-structured' { param($wb)
    $ws=$wb.Worksheets.Item('First half of the year'); $t=$ws.ListObjects.Item('Q1_Sales'); $r=$t.ListColumns.Item('Inactive').DataBodyRange; $r.Formula='=COUNTBLANK([@[January]:[June]])'; Release-Com $r; Release-Com $t; Release-Com $ws
}
Check 'T06' 'structured formula equivalent' $true $p 'CountBlankFormulaByHeaders' @('First half of the year','Q1_Sales','Inactive',$months)
$p = New-Case 't06-typed' { param($wb) $ws=$wb.Worksheets.Item('First half of the year'); $vals=@(0,2,0,0,0,2,1,1); for($i=0;$i -lt 8;$i++){$ws.Cells.Item(4+$i,9).Value2=$vals[$i]}; Release-Com $ws }
Check 'T06' 'typed outputs' $false $p 'CountBlankFormulaByHeaders' @('First half of the year','Q1_Sales','Inactive',$months)
$p = New-Case 't06-wrongrange' { param($wb) $ws=$wb.Worksheets.Item('First half of the year'); for($row=4;$row -le 11;$row++){$ws.Cells.Item($row,9).Formula=('=COUNTBLANK(B{0}:F{0})' -f $row)}; Release-Com $ws }
Check 'T06' 'wrong source range' $false $p 'CountBlankFormulaByHeaders' @('First half of the year','Q1_Sales','Inactive',$months)
$p = New-Case 't06-missing' { param($wb) $ws=$wb.Worksheets.Item('First half of the year'); for($row=4;$row -le 11;$row++){$ws.Cells.Item($row,9).Formula=('=COUNTBLANK(B{0}:G{0})' -f $row)}; $ws.Cells.Item(11,9).ClearContents(); Release-Com $ws }
Check 'T06' 'one row without formula' $false $p 'CountBlankFormulaByHeaders' @('First half of the year','Q1_Sales','Inactive',$months)
$p = New-Case 't06-count' { param($wb) $ws=$wb.Worksheets.Item('First half of the year'); for($row=4;$row-le11;$row++){$ws.Cells.Item($row,9).Formula=('=COUNT(B{0}:G{0})'-f$row)}; Release-Com $ws }
Check 'T06' 'COUNT instead of COUNTBLANK' $false $p 'CountBlankFormulaByHeaders' @('First half of the year','Q1_Sales','Inactive',$months)

$p = New-Case 't07-primary' { param($wb)
    $ws=$wb.Worksheets.Item('Top Toys Category'); $t=$ws.ListObjects.Item('Table4'); $s=$t.Sort; $s.SortFields.Clear(); [void]$s.SortFields.Add($t.ListColumns.Item('Toys').DataBodyRange,0,1); $s.Header=1; $s.Apply(); Release-Com $s; Release-Com $t; Release-Com $ws
}
Check 'T07' 'only primary sort field' $false $p 'TableMultiLevelSortStateEquals' @('Top Toys Category','Table4','A2:B12',$sortHeaders,$sortOrders)
$p = New-Case 't07-secondasc' { param($wb)
    $ws=$wb.Worksheets.Item('Top Toys Category'); $t=$ws.ListObjects.Item('Table4'); $s=$t.Sort; $s.SortFields.Clear(); [void]$s.SortFields.Add($t.ListColumns.Item('Toys').DataBodyRange,0,1); [void]$s.SortFields.Add($t.ListColumns.Item('Total Sales').DataBodyRange,0,1); $s.Header=1; $s.Apply(); Release-Com $s; Release-Com $t; Release-Com $ws
}
Check 'T07' 'secondary sort ascending' $false $p 'TableMultiLevelSortStateEquals' @('Top Toys Category','Table4','A2:B12',$sortHeaders,$sortOrders)
$p = New-Case 't07-reverse' { param($wb)
    $ws=$wb.Worksheets.Item('Top Toys Category'); $t=$ws.ListObjects.Item('Table4'); $s=$t.Sort; $s.SortFields.Clear(); [void]$s.SortFields.Add($t.ListColumns.Item('Total Sales').DataBodyRange,0,2); [void]$s.SortFields.Add($t.ListColumns.Item('Toys').DataBodyRange,0,1); $s.Header=1; $s.Apply(); Release-Com $s; Release-Com $t; Release-Com $ws
}
Check 'T07' 'sort field priority reversed' $false $p 'TableMultiLevelSortStateEquals' @('Top Toys Category','Table4','A2:B12',$sortHeaders,$sortOrders)

$p = New-Case 't08-layout2' { param($wb) $ws=$wb.Worksheets.Item('Last half of the year'); $co=$ws.ChartObjects().Item('Chart 1'); $co.Chart.ApplyLayout(2); Release-Com $co; Release-Com $ws }
Check 'T08' 'quick layout 2' $false $p 'ChartQuickLayoutEquals' @('Last half of the year','Chart 1',3,51,'A4:A11','B3:H3','B4:H11')
$p = New-Case 't08-wrongtype' { param($wb) $ws=$wb.Worksheets.Item('Last half of the year'); $co=$ws.ChartObjects().Item('Chart 1'); $co.Chart.ApplyLayout(3); $co.Chart.ChartType=52; Release-Com $co; Release-Com $ws }
Check 'T08' 'layout 3 wrong chart type' $false $p 'ChartQuickLayoutEquals' @('Last half of the year','Chart 1',3,51,'A4:A11','B3:H3','B4:H11')
$p = New-Case 't08-layout1' { param($wb) $ws=$wb.Worksheets.Item('Last half of the year'); $co=$ws.ChartObjects().Item('Chart 1'); $co.Chart.ApplyLayout(1); Release-Com $co; Release-Com $ws }
Check 'T08' 'quick layout 1' $false $p 'ChartQuickLayoutEquals' @('Last half of the year','Chart 1',3,51,'A4:A11','B3:H3','B4:H11')
$p = New-Case 't08-layout4' { param($wb) $ws=$wb.Worksheets.Item('Last half of the year'); $co=$ws.ChartObjects().Item('Chart 1'); $co.Chart.ApplyLayout(4); Release-Com $co; Release-Com $ws }
Check 'T08' 'quick layout 4' $false $p 'ChartQuickLayoutEquals' @('Last half of the year','Chart 1',3,51,'A4:A11','B3:H3','B4:H11')
$p = New-Case 't08-wrongchart' { param($wb) $ws=$wb.Worksheets.Item('First half of the year'); $co=$ws.ChartObjects().Item('Chart 1'); $co.Chart.ApplyLayout(3); Release-Com $co; Release-Com $ws }
Check 'T08' 'layout 3 on another chart only' $false $p 'ChartQuickLayoutEquals' @('Last half of the year','Chart 1',3,51,'A4:A11','B3:H3','B4:H11')

$results | Format-Table -AutoSize
$failed = @($results | Where-Object { -not $_.OK })
Write-Output ("TOTAL={0}; FAILED={1}" -f $results.Count,$failed.Count)
if($failed.Count -gt 0){ exit 1 }
