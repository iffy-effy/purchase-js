param(
  [Parameter(Mandatory=$true)][ValidateSet("clear","add","run","rows")][string]$Action,
  [string]$Item, $Qty, [string]$Due, [string]$Note = "", [string]$Vendor = "",
  [switch]$Save
)
# PR자동화 엑셀 매크로를 메시지창 없이 호출한다 (ME51N_PR.bas 의 PR_* 함수)
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$xl = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$wb = $null; foreach ($w in $xl.Workbooks) { if ($w.Name -like "ME51N_*PR자동화.xlsm") { $wb = $w } }
if (-not $wb) { "ERR|PR자동화 엑셀 파일이 열려 있지 않습니다"; exit 1 }
$pre = "'" + $wb.Name + "'!"
switch ($Action) {
  "clear" { $xl.Run($pre + "PR_ClearInput") }
  "add"   { if (-not $Due) { $Due = (Get-Date).ToString("yyyy-MM-dd") }
            $xl.Run($pre + "PR_AddRow", $Item, [double]$Qty, $Due, $Note, $Vendor) }
  "run"   { $xl.Run($pre + "PR_Run", [bool]$Save) }
  "rows"  { $ws = $wb.Sheets.Item("PR입력")
            foreach ($r in 5..54) { $b = [string]$ws.Cells.Item($r,2).Value2; if ($b) {
              "{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}" -f $r, $b, $ws.Cells.Item($r,3).Text, $ws.Cells.Item($r,4).Text, $ws.Cells.Item($r,5).Text,
                $ws.Cells.Item($r,6).Text, $ws.Cells.Item($r,11).Text, $ws.Cells.Item($r,12).Text, $ws.Cells.Item($r,15).Text, $ws.Cells.Item($r,16).Text } } }
}
