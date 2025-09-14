# 홈(현재) 폴더 구조 예:
# Home/
#   ├─ Session/
#   │    ├─ <AnySession>/{Light,Dark,Flat,Bias}  또는 {Light,masters(dark_stacked.fit,pp_flat_stacked.fit)}
#   │    └─ ...
#   └─ Bias/   (선택: 공통 Bias 사용시)

# run_all.ps1 ? 세션별 스택 + 프레임수 가중치 최종 합성
# (Siril 1.4 beta2 호환 / masters 폴더 자동 생성 및 dark_stacked.fit, pp_flat_stacked.fit 저장)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# === 사용자 설정 ===
$SirilCli = "C:\Program Files\Siril\bin\siril-cli.exe"   # siril-cli.exe 경로
$KeepSsf  = $false                                       # true면 SSF 보존, false면 마지막에 삭제

# === 기본 경로 ===
$HomeDir       = (Get-Location).Path
$StackPath = Join-Path $HomeDir "stack_all.fit"
$SessionRoot   = Join-Path $HomeDir "Session"
$CommonBiasDir = Join-Path $HomeDir "Bias"
$SessionStacks = Join-Path $HomeDir "_SESSION_STACKS"
New-Item -ItemType Directory -Force -Path $SessionStacks | Out-Null

# === 유틸 ===
function ToSirilPath([string]$p) { ($p -replace '\\','/') }
function Write-NoBom([string]$Path, [string]$Content) {
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
  Write-Host "  SSF saved: $Path"
}
function CountFrames([string]$Dir) {
  if (-not (Test-Path $Dir)) { return 0 }
  (Get-ChildItem $Dir -Recurse -File -Include *.fit,*.fits,*.xisf,*.tif,*.tiff -ErrorAction SilentlyContinue | Measure-Object).Count
}
function FlattenTo([string]$SrcDir, [string]$DstDir) {
  if (-not (Test-Path $SrcDir)) { return $false }
  New-Item -ItemType Directory -Force -Path $DstDir | Out-Null
  $files = Get-ChildItem $SrcDir -Recurse -File -Include *.fit,*.fits,*.xisf,*.tif,*.tiff -ErrorAction SilentlyContinue
  if ($files.Count -eq 0) { return $false }
  foreach ($f in $files) { Copy-Item $f.FullName -Dest (Join-Path $DstDir $f.Name) -Force }
  return $true
}
function Invoke-Siril([string]$SirilCli, [string]$SsfPath, [string]$LogPath) {
  $old = $ErrorActionPreference
  $ErrorActionPreference = 'Continue'   # Siril stderr로 인한 중단 방지
  try {
    $out = & "$SirilCli" -s "$SsfPath" 2>&1
    if ($LogPath) { $out | Out-File -FilePath $LogPath -Encoding utf8 }
    return ($out -join [Environment]::NewLine)
  } finally {
    $ErrorActionPreference = $old
  }
}

if (-not (Test-Path $SessionRoot)) { throw "Session 폴더가 없습니다: $SessionRoot" }

# === 공통 Bias 처리 ===
$CommonBiasFile = $null
$CreatedSsf = [System.Collections.Generic.List[string]]::new()

if (Test-Path $CommonBiasDir) {
  $cand = @("masterbias.fit","bias_stacked.fit") | ForEach-Object { Join-Path $CommonBiasDir $_ }
  $CommonBiasFile = $cand | Where-Object { Test-Path $_ } | Select-Object -First 1
  if ($CommonBiasFile) {
    Write-Host "▶ 공통 Bias: $(Split-Path $CommonBiasFile -Leaf) 사용"
  } elseif ((CountFrames $CommonBiasDir) -gt 0) {
    Write-Host "▶ 공통 Bias 생성"
    $ssf = @"
requires 1.2
cd $(ToSirilPath $CommonBiasDir)
convert .
stack . med -out=masterbias

"@
    $ssfPath = Join-Path $HomeDir "___make_common_bias.ssf"
    Write-NoBom $ssfPath $ssf
    $CreatedSsf.Add($ssfPath)
    & "$SirilCli" -s "$ssfPath"
    $CommonBiasFile = Join-Path $CommonBiasDir "masterbias.fit"
  }
}

# === 세션 루프 ===
$report = New-Object System.Collections.Generic.List[string]
$sessionDirs = Get-ChildItem $SessionRoot -Directory | Sort-Object Name
if ($sessionDirs.Count -eq 0) { throw "Session 하위에 세션 폴더가 없습니다." }

foreach ($sess in $sessionDirs) {
  $SPath = $sess.FullName; $SName = $sess.Name
  $LightDir   = Join-Path $SPath "Light"
  $DarkDir    = Join-Path $SPath "Dark"
  $FlatDir    = Join-Path $SPath "Flat"
  $BiasDir    = Join-Path $SPath "Bias"
  $MastersDir = Join-Path $SPath "masters"

  # masters 디렉토리 보장 (항상 존재)
  New-Item -ItemType Directory -Force -Path $MastersDir | Out-Null

  # _WORK 평탄화
  $WorkDir = Join-Path $SPath "_WORK"
  $W_LIGHT = Join-Path $WorkDir "LIGHT"
  $W_DARK  = Join-Path $WorkDir "DARK"
  $W_FLAT  = Join-Path $WorkDir "FLAT"
  $W_BIAS  = Join-Path $WorkDir "BIAS"
  New-Item -ItemType Directory -Force -Path $WorkDir | Out-Null

  $hasLight = FlattenTo $LightDir $W_LIGHT
  if (-not $hasLight) { Write-Warning "세션 '$SName' → Light 없음, 건너뜀"; continue }
  $hasDark  = FlattenTo $DarkDir  $W_DARK
  $hasFlat  = FlattenTo $FlatDir  $W_FLAT
  $hasBias  = FlattenTo $BiasDir  $W_BIAS

# masters 폴더 보장
if (-not (Test-Path $MastersDir)) {
  New-Item -ItemType Directory -Force -Path $MastersDir | Out-Null
}

  # masters 재사용 탐색
  $MasterDark = $null; $MasterFlat = $null
  $candBias = @("bias_stacked.fit","masterbias.fit") | ForEach-Object { Join-Path $MastersDir $_ }
  $candDark = @("dark_stacked.fit","masterdark.fit") | ForEach-Object { Join-Path $MastersDir $_ }
  $candFlat = @("pp_flat_stacked.fit","masterflat.fit") | ForEach-Object { Join-Path $MastersDir $_ }
  $MasterBias = $candBias | Where-Object { Test-Path $_ } | Select-Object -First 1
  $MasterDark = $candDark | Where-Object { Test-Path $_ } | Select-Object -First 1
  $MasterFlat = $candFlat | Where-Object { Test-Path $_ } | Select-Object -First 1

	# Flat용 Bias 결정 (드롭인 교체)
	$BiasArgForFlat = ""
	$BiasBlock      = ""
	$BiasStatus     = "없음"   # 재사용 | 생성 | 공통 | 없음

	if ($MasterBias) {
	  # masters에 masterbias/bias_stacked가 이미 있음 → 재사용
	  $BiasArgForFlat = " -bias=$(ToSirilPath $MasterBias)"
	  $BiasStatus     = "재사용"

	} elseif ($hasBias -and (CountFrames $W_BIAS) -gt 0) {
	  # 세션 BIAS 프레임으로 마스터 생성
	  $BiasBlock = @"
cd $(ToSirilPath $W_BIAS)
convert .
stack . med -out=$(ToSirilPath (Join-Path $MastersDir 'bias_stacked'))

"@
	  $MasterBias     = Join-Path $MastersDir 'bias_stacked.fit'
	  $BiasArgForFlat = " -bias=$(ToSirilPath $MasterBias)"
	  $BiasStatus     = "생성"

	} elseif ($CommonBiasFile) {
	  # 공통 Bias 사용
	  $BiasArgForFlat = " -bias=$(ToSirilPath $CommonBiasFile)"
	  $BiasStatus     = "공통"

	} else {
	  # 정말로 Bias가 없음 → Flat 보정 시 Bias 미사용
	  $BiasStatus = "없음"
	}

$DarkStatus = if ($MasterDark) { '재사용' } elseif ($hasDark) { '생성' } else { '없음' }
$FlatStatus = if ($MasterFlat) { '재사용' } elseif ($hasFlat) { '생성' } else { '없음' }
# $BiasStatus 는 이미 위에서 계산됨

Write-Host "▶ 세션: $SName  (Dark:$DarkStatus, Flat:$FlatStatus, Bias:$BiasStatus)"

$ssf = @"
requires 1.2
$BiasBlock

"@

# 1) Dark master
if (-not $MasterDark) {
  if ($hasDark -and (CountFrames $W_DARK) -gt 0) {
    $ssf += @"
cd $(ToSirilPath $W_DARK)
convert .
stack ._ med -nonorm -out=$(ToSirilPath (Join-Path $MastersDir "dark_stacked"))

"@  # ← 빈 줄 유지 (개행 보장)
    $MasterDark = Join-Path $MastersDir "dark_stacked.fit"
  } else {
    Write-Warning "세션 '$SName' → Dark 없음: 라이트 보정 시 Dark 미사용"
  }
}

# 2) Flat master
if (-not $MasterFlat) {
  if ($hasFlat -and (CountFrames $W_FLAT) -gt 0) {
    $ssf += @"
cd $(ToSirilPath $W_FLAT)
convert .
calibrate ._$BiasArgForFlat -cfa -prefix=cb_flat_
stack cb_flat_._ med -norm=mul -out=$(ToSirilPath (Join-Path $MastersDir "pp_flat_stacked"))

"@  # ← 빈 줄 유지
    $MasterFlat = Join-Path $MastersDir "pp_flat_stacked.fit"
  } else {
    Write-Warning "세션 '$SName' → Flat 없음: 라이트 보정 시 Flat 미사용"
  }
}

# 3) Light 처리
$calDark = ""; if ($MasterDark) { $calDark = " -dark=$(ToSirilPath $MasterDark)" }
$calFlat = ""; if ($MasterFlat) { $calFlat = " -flat=$(ToSirilPath $MasterFlat)" }

$ssf += @"
cd $(ToSirilPath $W_LIGHT)
convert .
calibrate ._$calDark$calFlat -cc=dark -cfa -equalize_cfa -debayer -prefix=pp_
register pp_._
stack r_pp_._ rej 3 3 -norm=addscale -output_norm -rgb_equal -32b -out=stack_session

"@  # ← 빈 줄 유지

  # 실행
  $ssfPath = Join-Path $SPath "___session_$SName.ssf"
  Write-NoBom $ssfPath $ssf
  $CreatedSsf.Add($ssfPath)

  $sessionLog = Join-Path $SPath "siril_session.log"
  $out = Invoke-Siril $SirilCli $ssfPath $sessionLog

  # 프레임 수 추출 (로그 → Integration of N images)
  $n = $null
  $rx = [regex]'Integration of\s+(\d+)\s+images'
  $matches = $rx.Matches($out)
  if ($matches.Count -gt 0) { $n = [int]$matches[$matches.Count-1].Groups[1].Value }
  if (-not $n) {
    $n = (Get-ChildItem -Path $W_LIGHT -Filter 'r_pp_._*.fit' -File -ErrorAction SilentlyContinue |
          Measure-Object | Select-Object -ExpandProperty Count)
  }
  $report.Add("$($SName): $n frames")

  # 세션 스택 결과 복사
  $sessStack = Join-Path $W_LIGHT "stack_session.fit"
  if (Test-Path $sessStack) {
    $dest = Join-Path $SessionStacks ("stack_{0}.fit" -f $SName)
    Copy-Item $sessStack $dest -Force
  }
}

# === 최종 합성 ===
$stackFiles = Get-ChildItem (Join-Path $SessionStacks "stack_*.fit") -ErrorAction SilentlyContinue
if ($stackFiles.Count -eq 0) { throw "세션 스택을 하나도 만들지 못했습니다." }

$finalSsf = @"
requires 1.2
cd $(ToSirilPath $SessionStacks)
convert .
register ._
stack r_._ rej 3 3 -norm=addscale -output_norm -rgb_equal -32b -weight=nbstack -out=stack_all

# === 후처리 ===
load stack_all

# 원본 선형 스택 보존
save $(ToSirilPath $StackPath)

########### 일반처리

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

$finalPath = Join-Path $HomeDir "___final_weighted_stack.ssf"
Write-NoBom $finalPath $finalSsf
$CreatedSsf.Add($finalPath)

Write-Host "▶ 최종 가중치 스택 실행..."
$finalLog = Join-Path $HomeDir "final_siril.log"
$finalOut = Invoke-Siril $SirilCli $finalPath $finalLog
Write-Host "  · 최종 스택 로그: $finalLog"

# === 리포트 ===
$report.Insert(0, ("총 세션 스택: {0}개" -f $stackFiles.Count))
$reportPath = Join-Path $HomeDir "session_stack_report.txt"
[System.IO.File]::WriteAllLines($reportPath, $report)

# === 정리: WORK + 중간산출물 삭제 ===
Write-Host "▶ 중간 산출물 정리..."
# _WORK 폴더 삭제
Get-ChildItem -Path (Join-Path $SessionRoot '*') -Directory -Recurse |
  Where-Object { $_.Name -eq '_WORK' } |
  Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

# 홈 아래 중간 산출물 삭제
Get-ChildItem -Path (Join-Path $HomeDir '*') -File -Recurse -ErrorAction SilentlyContinue |
  Where-Object { $_.Name -like 'r_all_*.fit' -or $_.Name -like '*.seq' -or $_.Name -like '._*.fit' -or $_.Name -like 'pp_._*.fit' -or $_.Name -like 'r_pp_._*.fit' } |
  Remove-Item -Force

# === 생성한 SSF 자동 삭제(옵션) ===
if (-not $KeepSsf) {
  foreach ($p in $CreatedSsf) {
    if (Test-Path $p) { Remove-Item $p -Force -ErrorAction SilentlyContinue }
  }
}

Write-Host "`n? 완료! 최종 파일: $(Join-Path $HomeDir 'stack_all.fit')"
Write-Host "?? 세션별 프레임 수 요약: $reportPath"
