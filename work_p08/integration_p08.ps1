$ErrorActionPreference='Stop'
$repo='C:\Users\HUYNH HAU\source\repos'
$starter=Join-Path $repo 'MosTrainer\Projects\Excel2019_P08\starter.xlsx'
$path=Join-Path $env:TEMP 'MosTrainer-P08-integration-all.xlsx'
Copy-Item -LiteralPath $starter -Destination $path -Force
function RC($o){if($null-ne$o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}}
$x=$null;$w=$null
try{
  $x=New-Object -ComObject Excel.Application;$x.Visible=$false;$x.DisplayAlerts=$false;$w=$x.Workbooks.Open($path)
  $s=$w.Worksheets.Item('Forecast');$t=$s.ListObjects.Item('Table5');for($r=1;$r-le$t.DataBodyRange.Rows.Count;$r++){$row=$t.DataBodyRange.Row+$r-1;$s.Cells.Item($row,3).Formula=('=B{0}*Q2_Increase'-f$row)};RC $t;RC $s
  $s=$w.Worksheets.Item('Suppliers');$t=$s.ListObjects.Item('Table2');$t.ListRows.Item(2).Delete();RC $t;RC $s
  $s=$w.Worksheets.Item('First half of the year');$r=$s.Range('A4:A11');$r.HorizontalAlignment=-4131;$r.IndentLevel=1;RC $r
  $r=$s.Range('J4:J11');[void]$r.SparklineGroups.Add(3,"'First half of the year'!B4:G11");RC $r
  $t=$s.ListObjects.Item('Q1_Sales');$t.ShowTotals=$true;foreach($h in @('January','February','March','April','May','June','Total')){$t.ListColumns.Item($h).TotalsCalculation=1};RC $t
  for($row=4;$row-le11;$row++){$s.Cells.Item($row,9).Formula=('=COUNTBLANK(B{0}:G{0})'-f$row)};RC $s
  $s=$w.Worksheets.Item('Top Toys Category');$t=$s.ListObjects.Item('Table4');$sort=$t.Sort;$sort.SortFields.Clear();[void]$sort.SortFields.Add($t.ListColumns.Item('Toys').DataBodyRange,0,1);[void]$sort.SortFields.Add($t.ListColumns.Item('Total Sales').DataBodyRange,0,2);$sort.Header=1;$sort.Apply();RC $sort;RC $t;RC $s
  $s=$w.Worksheets.Item('Last half of the year');$co=$s.ChartObjects().Item('Chart 1');$co.Chart.ApplyLayout(3);RC $co;RC $s
  $w.Save()
}finally{if($null-ne$w){$w.Close($false)};if($null-ne$x){$x.Quit()};RC $w;RC $x;[GC]::Collect();[GC]::WaitForPendingFinalizers()}

[void][Reflection.Assembly]::LoadFrom((Join-Path $repo 'MosTrainer.Core\bin\Debug\MosTrainer.Core.dll'))
[void][Reflection.Assembly]::LoadFrom((Join-Path $repo 'MosTrainer.Projects\bin\Debug\Newtonsoft.Json.dll'))
[void][Reflection.Assembly]::LoadFrom((Join-Path $repo 'MosTrainer.Projects\bin\Debug\MosTrainer.Projects.dll'))
[void][Reflection.Assembly]::LoadFrom((Join-Path $repo 'MosTrainer.Excel\bin\Debug\MosTrainer.Excel.dll'))
$validator=New-Object MosTrainer.Projects.ProjectValidator
$validated=$validator.Validate((Join-Path $repo 'MosTrainer\Projects\Excel2019_P08'),'vi')
$controller=New-Object MosTrainer.Excel.ExcelController
try{
  $controller.OpenWorkbook($path);$service=New-Object MosTrainer.Core.Services.GradingService($controller)
  $failed=0
  foreach($task in $validated.Tasks){$task.ProjectId='Excel2019_P08';$result=$service.CheckTask($task);$ok=[bool]$result.Item1;if(-not$ok){$failed++};[pscustomobject]@{Task=$task.TaskId;Assertion=$task.AssertionType;Pass=$ok;Message=$result.Item2}}
  Write-Output ("ROUTE_TOTAL={0}; ROUTE_FAILED={1}"-f$validated.Tasks.Count,$failed)
  if($failed-gt0){exit 1}
}finally{$controller.Dispose()}
