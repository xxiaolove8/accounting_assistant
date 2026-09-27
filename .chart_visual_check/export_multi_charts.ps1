param([string]$Suffix = '', [ValidateSet('all', 'bar', 'line', 'pie')][string]$Kind = 'all')
$ErrorActionPreference = 'Stop'
$checkDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $checkKinds = if ($Kind -eq 'all') { @('bar', 'line', 'pie') } else { @($Kind) }
  foreach ($kind in $checkKinds) {
    $sourcePath = Join-Path $checkDir "multi_$kind.xlsx"
    $book = $excel.Workbooks.Open($sourcePath, 0, $true)
    try {
      $sheet = $book.Worksheets.Item(1)
      $charts = $sheet.ChartObjects()
      for ($index = 1; $index -le $charts.Count; $index++) {
        $chart = $charts.Item($index).Chart
        $targetPath = Join-Path $checkDir "multi_${kind}_${index}${Suffix}.png"
        if (-not $chart.Export($targetPath, 'PNG')) { throw "Chart export failed: $targetPath" }
        $seriesNames = @()
        for ($seriesIndex = 1; $seriesIndex -le $chart.SeriesCollection().Count; $seriesIndex++) {
          $seriesNames += $chart.SeriesCollection($seriesIndex).Name
        }
        Write-Output "$kind chart $index series: $($seriesNames -join ', ')"
        if ($kind -ne 'pie') {
          Write-Output "axes: $($chart.Axes(1).AxisTitle.Text) / $($chart.Axes(2).AxisTitle.Text)"
        }
      }
    } finally {
      $book.Close($false)
    }
  }
} finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
}
