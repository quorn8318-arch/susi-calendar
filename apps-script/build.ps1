# 구글 Apps Script 웹 앱용 파일 만들기
# 사용: powershell -ExecutionPolicy Bypass -File apps-script\build.ps1
# 결과: apps-script\Index.html (앱 화면·매칭 코드·달력 틀·대학 일정 목록을 한 파일로 합친 것), apps-script\Code.gs
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$utf8 = New-Object System.Text.UTF8Encoding($false)
$index    = [IO.File]::ReadAllText((Join-Path $root 'index.html'), $utf8)
$match    = [IO.File]::ReadAllText((Join-Path $root 'match.js'), $utf8)
$template = [IO.File]::ReadAllText((Join-Path $root 'calendar-template.html'), $utf8)
$sched    = [Convert]::ToBase64String([IO.File]::ReadAllBytes((Join-Path $root 'schedule-2027.xlsx')))

# 달력 틀을 자바스크립트 문자열로 (</script>가 페이지를 끊지 않게 '<'를 <로)
Add-Type -AssemblyName System.Web.Extensions
$ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$ser.MaxJsonLength = [int]::MaxValue
$tplJson = $ser.Serialize($template).Replace('<', '<')
$embed = "<script>window.SUSI_EMBED = { template: $tplJson, schedule: `"$sched`" };</script>"

$tag = '<script src="match.js"></script>'
if (-not $index.Contains($tag)) { throw 'index.html에서 match.js 태그를 찾지 못했습니다.' }
$out = $index.Replace($tag, "$embed`n<script>`n$match`n</script>")
[IO.File]::WriteAllText((Join-Path $PSScriptRoot 'Index.html'), $out, $utf8)
"Index.html: {0:N0} bytes" -f $utf8.GetByteCount($out)
