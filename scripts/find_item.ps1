param([Parameter(Mandatory=$true)][string]$Query, [int]$Top = 15)
# 품목마스터에서 품목을 찾는다 (엑셀 매크로 ResolveItemCell과 같은 규칙)
# 출력: 줄마다  번호|품목명|자재코드|단위|요청횟수|구분|최근요청사유
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$xl = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$wb = $null; foreach ($w in $xl.Workbooks) { if ($w.Name -like "ME51N_*PR자동화.xlsm") { $wb = $w } }
if (-not $wb) { "ERR|PR자동화 엑셀 파일이 열려 있지 않습니다"; exit 1 }
$m = $wb.Sheets.Item("품목마스터")
$last = $m.Cells($m.Rows.Count, 2).End(-4162).Row
$v = $m.Range("B5:S$last").Value2   # 1=품목명 2=자재코드 3=내역 4=단위 14=구분 16=요청횟수 17=최근사유 18=별칭

function Norm([string]$s) { $s = $s.ToLower(); foreach ($ch in @(' ','(',')','[',']','-','_','*',',','.','/','&',[char]160)) { $s = $s.Replace([string]$ch,'') }; $s }

$q = $Query.Trim()
$rows = @()
$n = $last - 4
# 1) 품목명 정확히 일치 / 자재코드 일치
for ($i = 1; $i -le $n; $i++) {
  if ([string]$v[$i,1] -eq $q -or ([string]$v[$i,2] -ne "" -and [string]$v[$i,2] -eq $q)) {
    "1|{0}|{1}|{2}|{3}|{4}|{5}" -f $v[$i,1], $v[$i,2], $v[$i,4], $v[$i,16], $v[$i,14], $v[$i,17]; exit 0
  }
}
# 2) 띄어쓴 단어가 모두 포함된 품목
$tokens = $q -split '\s+' | Where-Object { $_ } | ForEach-Object { Norm $_ } | Where-Object { $_ }
for ($i = 1; $i -le $n; $i++) {
  if (-not $v[$i,1]) { continue }
  if (-not ([string]$v[$i,2]) -and -not ([string]$v[$i,3])) { continue }
  $hay = Norm ("{0}|{1}|{2}|{3}" -f $v[$i,1], $v[$i,2], $v[$i,3], $v[$i,18])
  $ok = $true; foreach ($t in $tokens) { if (-not $hay.Contains($t)) { $ok = $false; break } }
  if ($ok) { $rows += [pscustomobject]@{ name=[string]$v[$i,1]; code=[string]$v[$i,2]; unit=[string]$v[$i,4]; cnt=[double]("0"+[string]$v[$i,16]); cat=[string]$v[$i,14]; why=[string]$v[$i,17] } }
}
if ($rows.Count -eq 0) { "0|검색 결과 없음: $q"; exit 0 }
"#{0}건" -f $rows.Count
$k = 0
foreach ($r in ($rows | Sort-Object cnt -Descending | Select-Object -First $Top)) {
  $k++; "{0}|{1}|{2}|{3}|{4}|{5}|{6}" -f $k, $r.name, $r.code, $r.unit, $r.cnt, $r.cat, $r.why
}
