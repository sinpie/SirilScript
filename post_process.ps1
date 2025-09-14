# 후처리: stack_all.fit → HOO 합성 + 오토스트레치 + 배경추출 + asinh + JPEG
# 사용법:  PowerShell에서 이 파일이 있는 폴더에서 실행
#   PS> .\postprocess_stack.ps1

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- 설정 ---
$SirilCli  = "C:\Program Files\Siril\bin\siril-cli.exe"
$HomeDir   = (Get-Location).Path
$StackPath = Join-Path $HomeDir "stack_all.fit"     # 입력
$KeepSsf   = $false                                 # 생성 SSF 보관 여부

function ToSirilPath([string]$p) { ($p -replace '\\','/') }
if (!(Test-Path $StackPath)) { throw "입력 파일이 없습니다: $StackPath" }

# --- _WORK 폴더 생성 ---
$WorkDir = Join-Path $HomeDir "_WORK"
New-Item -ItemType Directory -Force -Path $WorkDir | Out-Null

# --- SSF 생성 ---
$ssf = @"
requires 1.2
# 작업 폴더
cd $(ToSirilPath $WorkDir)

########### 일반처리
load $(ToSirilPath $StackPath)

# 1) Histogram Auto Stretch
autostretch -linked

# 2) Background Extraction
subsky -rbf -samples=20 -smooth=0.5 -tolerance=2.0

# 3) Asinh Transform (blackpoint 0.2)
asinh 1 0.2 -clipmode=globalrescale

# 4) 결과 저장 (중간 FIT/JPG는 _WORK, 최종 FIT/JPG는 홈)
save stack_all_processed
savejpg stack_all_final 95
save $(ToSirilPath (Join-Path $HomeDir "stack_all_processed"))
savejpg $(ToSirilPath (Join-Path $HomeDir "stack_all_final")) 95

########### HOO처리
load $(ToSirilPath $StackPath)

# 1) 채널 분리 -> _WORK에 R/G/B 파일 생성
split stack_R stack_G stack_B

# 2) OIII 합성 (G 크게, B 작게)
pm '`$stack_G$ * 0.75 + `$stack_B$ * 0.25' -nosum
save OIII_lin

# 3) OIII 부스트 (필요에 따라 계수 2.0~3.0 조정)
pm '`$OIII_lin$ * 2.5' -nosum
save OIII_boost

# 4) HOO 합성: R=stack_R, G=OIII, B=OIII
rgbcomp stack_R OIII_boost OIII_boost
load composed_rgb

# 5) 오토 스트레치
autostretch -linked

# 6) Background Extraction (다항식 근사)
subsky -rbf -samples=20 -smooth=0.5 -tolerance=2.0

# 7) Asinh 변환
asinh 1 0.2 -clipmode=globalrescale

# 8) JPEG로 저장 (_WORK 중간본 + 홈 최종본)
save stack_all_processed_HOO
savejpg stack_all_final_HOO 95
save $(ToSirilPath (Join-Path $HomeDir "stack_all_processed_HOO"))
savejpg $(ToSirilPath (Join-Path $HomeDir "stack_all_final_HOO")) 95

"@

$ssfPath = Join-Path $WorkDir "___postprocess_stack_all.ssf"
[IO.File]::WriteAllText($ssfPath, $ssf, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "SSF saved: $ssfPath"

# --- 실행 & 로그 저장 ---
$logPath = Join-Path $HomeDir "postprocess_siril.log"
$old = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
try {
  $out = & "$SirilCli" -s "$ssfPath" 2>&1
  $out | Out-File -FilePath $logPath -Encoding utf8
  Write-Host "로그 저장: $logPath"
} finally { $ErrorActionPreference = $old }

# --- 정리(옵션) ---
if (-not $KeepSsf -and (Test-Path $ssfPath)) {
  Remove-Item $ssfPath -Force -ErrorAction SilentlyContinue
}

Write-Host "`n? 완료!"
Write-Host " - 최종 FIT  : $(Join-Path $HomeDir 'stack_all_HOO.fit')"
Write-Host " - 최종 JPEG : $(Join-Path $HomeDir 'stack_all_final_HOO.jpg')"
Write-Host " - 중간 산출물: $WorkDir"
