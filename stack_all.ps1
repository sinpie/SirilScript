# 20260216 DarkFlat추가
# 20260218 세션별 pp_light (보정된 light)를 한번에 스택하도록 변경

<# 
Home/
  ├─ Session/
  │    ├─ <AnySession>/{Light,Dark,DarkFlat,Flat,Bias}  또는 {Light,masters(...)}
  │    └─ ...
  ├─ Bias/   (선택: 공통 Bias)
  └─ Dark/   (선택: 공통 Dark)
공통 DarkFlat은 없음(세션 Flat 기반)

run_all.ps1 : 세션별 보정+등록 -> _ALL_REGISTERED에 모아 최종 스택
Siril 1.4.x (requires 1.2)
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# =========================
# 사용자 설정
# =========================
$SirilCli = "C:\Program Files\Siril\bin\siril-cli.exe"
$KeepSsf  = $false

# =========================
# 상수(타입 키)
# =========================
$T_BIAS     = "Bias"
$T_DARK     = "Dark"
$T_DARKFLAT = "DarkFlat"
$T_FLAT     = "Flat"
$T_LIGHT    = "Light"

# =========================
# 기본 경로
# =========================
$HomeDir       = (Get-Location).Path
$SessionRoot   = Join-Path $HomeDir "Session"
$CommonBiasDir = Join-Path $HomeDir "Bias"
$CommonDarkDir = Join-Path $HomeDir "Dark"
$AllRegistered = Join-Path $HomeDir "_ALL_REGISTERED"
$StackPath     = Join-Path $HomeDir "stack_all.fit"

New-Item -ItemType Directory -Force -Path $AllRegistered | Out-Null

# =========================
# 유틸
# =========================
function ToSirilPath([string]$p) { (($p -replace '\\','/').Trim()) }
function ToSirilQ([string]$p) {
  if ([string]::IsNullOrWhiteSpace($p)) { return '""' }
  $s = (ToSirilPath $p).Trim().Trim('"')
  return '"' + $s + '"'
}

function Write-NoBom([string]$Path, [string]$Content) {
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
  Write-Host "  SSF saved: $Path"
}

function Invoke-Siril([string]$SirilCli, [string]$SsfPath, [string]$LogPath) {
  $old = $ErrorActionPreference
  $ErrorActionPreference = 'Continue'  # Siril stderr로 중단 방지
  try {
    $out = & "$SirilCli" -s "$SsfPath" 2>&1
    if ($LogPath) { $out | Out-File -FilePath $LogPath -Encoding utf8 }
    return ($out -join [Environment]::NewLine)
  } finally {
    $ErrorActionPreference = $old
  }
}

function Find-Master([string]$Dir, [string[]]$names) {
  $names | ForEach-Object { Join-Path $Dir $_ } |
    Where-Object { Test-Path $_ } |
    Select-Object -First 1
}

function CountFrames([string]$Dir) {
  if (-not (Test-Path $Dir)) { return 0 }
  (Get-ChildItem $Dir -Recurse -File -Include *.fit,*.fits,*.fts,*.xisf,*.tif,*.tiff -ErrorAction SilentlyContinue |
    Measure-Object).Count
}

# "필요할 때만 복사"를 위해: 복사하면서 개수 리턴
function FlattenToCount([string]$SrcDir, [string]$DstDir) {
  if (-not (Test-Path $SrcDir)) { return 0 }
  $files = Get-ChildItem $SrcDir -Recurse -File -Include *.fit,*.fits,*.fts,*.xisf,*.tif,*.tiff -ErrorAction SilentlyContinue
  if (-not $files -or $files.Count -eq 0) { return 0 }

  New-Item -ItemType Directory -Force -Path $DstDir | Out-Null

  foreach ($f in $files) {
    Copy-Item $f.FullName -Destination (Join-Path $DstDir $f.Name) -Force
  }
  return $files.Count
}

# Siril 옵션 인자용 상대경로(대상 파일 존재여부 무관)
function RelPath([string]$FromDir, [string]$ToPath) {
  if (-not (Test-Path -LiteralPath $FromDir)) {
    New-Item -ItemType Directory -Force -Path $FromDir | Out-Null
  }
  $fromFull = [System.IO.Path]::GetFullPath($FromDir)

  $toFull = if (Test-Path -LiteralPath $ToPath) {
    (Resolve-Path -LiteralPath $ToPath).Path
  } else {
    [System.IO.Path]::GetFullPath($ToPath)
  }

  $fromUri = [Uri]($fromFull.TrimEnd('\') + '\')
  $toUri   = [Uri]$toFull
  $rel = $fromUri.MakeRelativeUri($toUri).ToString().Replace('/', '\')
  return $rel
}

# =========================
# Plan/State
# =========================
function New-CalState([string]$TypeKey) {
  [pscustomobject]@{
    Type    = $TypeKey
    SrcDir  = $null
    WorkDir = $null

    Copied  = 0        # _WORK로 복사된 프레임 수
    Master  = $null    # 사용할 master 파일 경로(세션 masters 또는 공통)
    Create  = $false   # master 생성 필요 여부

    Status  = "없음"   # 재사용/공통/생성/없음
  }
}

function New-Plan([string]$SPath, [string]$SName, [string]$HomeDir, [string]$CommonBias, [string]$CommonDark) {
  $masters = Join-Path $SPath "masters"
  $work    = Join-Path $SPath "_WORK"

  [pscustomobject]@{
    HomeDir     = $HomeDir
    SPath       = $SPath
    SName       = $SName

    MastersDir  = $masters
    WorkRoot    = $work

    CommonBias  = $CommonBias   # full path or $null
    CommonDark  = $CommonDark   # full path or $null

    Cal         = @{}           # TypeKey -> CalState
  }
}

# "masters 재사용 / 공통 사용 / 세션프레임 있으면 생성" 공통 로직
function Resolve-FromSessionOrCommon {
  param(
    [pscustomobject]$Plan,
    [pscustomobject]$State,
    [string[]]$ReuseNames,
    [string]$CommonMasterPath,   # Bias/Dark만 사용(없으면 $null)
    [string]$CreateMasterName    # 생성 시 masters에 저장할 파일명
  )

  # 1) masters 재사용
  $reuse = Find-Master $Plan.MastersDir $ReuseNames
  if ($reuse) {
    $State.Master = $reuse
    $State.Status = "세션"
    return
  }

  # 2) 공통 master
  if ($CommonMasterPath) {
    $State.Master = $CommonMasterPath
    $State.Status = "공통"
    return
  }

  # 3) 세션 프레임 있으면 생성 예정
  if ($State.Copied -gt 0) {
    $State.Master = Join-Path $Plan.MastersDir $CreateMasterName
    $State.Create = $true
    $State.Status = "생성"
    return
  }

  $State.Master = $null
  $State.Status = "없음"
}

# =========================
# 공통 마스터 생성( Bias / Dark )
# =========================
function Ensure-CommonMaster {
  param(
    [ValidateSet("bias","dark")] [string]$Kind,
    [string]$CommonDir,
    [string]$HomeDir,
    [string]$SirilCli,
    [string]$CommonBiasForDark  # dark일 때만 사용(없으면 $null)
  )

  if (-not (Test-Path $CommonDir)) { return $null }

  $reuseNames = if ($Kind -eq "bias") { @("masterbias.fit","bias_stacked.fit") } else { @("masterdark.fit","dark_stacked.fit") }
  $cand = $reuseNames | ForEach-Object { Join-Path $CommonDir $_ }
  $reuse = $cand | Where-Object { Test-Path $_ } | Select-Object -First 1
  if ($reuse) { return (Resolve-Path $reuse).Path }

  if ((CountFrames $CommonDir) -le 0) { return $null }

  Write-Host "▶ 공통 $Kind 생성"
  
  $calLine = ""
  $outName = if ($Kind -eq "bias") { "masterbias" } else { "masterdark" }
  $stackLine = if ($Kind -eq "bias") { "stack bias_ med -out=$outName" } else { "stack dark_ med -nonorm -out=$outName" }

<#
  $outName = if ($Kind -eq "bias") { "masterbias" } else { "masterdark" }
  $stackLine = if ($Kind -eq "bias") { "stack bias_ med -out=$outName" } else { "stack pp_dark_ med -nonorm -out=$outName" }

  # dark는 bias가 있으면 제거 후 stack
  $calLine = ""
  if ($Kind -eq "dark") {
    if ($CommonBiasForDark) {
      $calLine = "calibrate dark_ -bias=$(RelPath $CommonDir $CommonBiasForDark) -prefix=pp_"
    } else {
      # bias 없으면 그냥 pp_ 없이 stack해도 되지만, 뒤 일관성을 위해 pp_를 안 만들거면 stack 대상도 바꿔야 함
      # -> bias 없으면 그냥 dark_를 stack
      $stackLine = "stack dark_ med -nonorm -out=$outName"
    }
  }
#>
  $ssf = @"
requires 1.2
cd $(ToSirilQ $CommonDir)
convert $Kind
$calLine
$stackLine

"@

  $ssfPath = Join-Path $HomeDir "___make_common_$Kind.ssf"
  Write-NoBom $ssfPath $ssf
  $null = Invoke-Siril $SirilCli $ssfPath $null

  $masterFit = Join-Path $CommonDir ($outName + ".fit")
  if (Test-Path $masterFit) { return (Resolve-Path $masterFit).Path }

  Write-Warning "공통 $Kind 생성 후 $($outName).fit을 찾지 못했습니다. ($CommonDir)"
  return $null
}

# =========================
# 타입별 Resolve
# =========================
function Resolve-Bias {
  param([pscustomobject]$Plan)

  $s = New-CalState $T_BIAS
  $s.SrcDir  = Join-Path $Plan.SPath "Bias"
  $s.WorkDir = Join-Path $Plan.WorkRoot "BIAS"

  # masters 재사용이면 복사 불필요, 없으면 복사
  $reuse = Find-Master $Plan.MastersDir @("bias_stacked.fit","masterbias.fit")
  if (-not $reuse) { $s.Copied = FlattenToCount $s.SrcDir $s.WorkDir }

  # Bias는 공통 Bias 허용
  Resolve-FromSessionOrCommon -Plan $Plan -State $s `
    -ReuseNames @("bias_stacked.fit","masterbias.fit") `
    -CommonMasterPath $Plan.CommonBias `
    -CreateMasterName "bias_stacked.fit"

  $Plan.Cal[$T_BIAS] = $s
}

function Resolve-Dark {
  param([pscustomobject]$Plan)

  $s = New-CalState $T_DARK
  $s.SrcDir  = Join-Path $Plan.SPath "Dark"
  $s.WorkDir = Join-Path $Plan.WorkRoot "DARK"

  $reuse = Find-Master $Plan.MastersDir @("dark_stacked.fit","masterdark.fit")
  if (-not $reuse) { $s.Copied = FlattenToCount $s.SrcDir $s.WorkDir }

  # Dark는 공통 Dark 허용
  Resolve-FromSessionOrCommon -Plan $Plan -State $s `
    -ReuseNames @("dark_stacked.fit","masterdark.fit") `
    -CommonMasterPath $Plan.CommonDark `
    -CreateMasterName "dark_stacked.fit"

  $Plan.Cal[$T_DARK] = $s
}

function Resolve-DarkFlat {
  param([pscustomobject]$Plan)

  $s = New-CalState $T_DARKFLAT
  $s.SrcDir  = Join-Path $Plan.SPath "DarkFlat"
  $s.WorkDir = Join-Path $Plan.WorkRoot "DARKFLAT"

  # DarkFlat은 공통 없음
  $reuse = Find-Master $Plan.MastersDir @("darkflat_stacked.fit","masterdarkflat.fit")
  if (-not $reuse) { $s.Copied = FlattenToCount $s.SrcDir $s.WorkDir }

  Resolve-FromSessionOrCommon -Plan $Plan -State $s `
    -ReuseNames @("darkflat_stacked.fit","masterdarkflat.fit") `
    -CommonMasterPath $null `
    -CreateMasterName "darkflat_stacked.fit"

  $Plan.Cal[$T_DARKFLAT] = $s
}

function Resolve-Flat {
  param([pscustomobject]$Plan)

  $s = New-CalState $T_FLAT
  $s.SrcDir  = Join-Path $Plan.SPath "Flat"
  $s.WorkDir = Join-Path $Plan.WorkRoot "FLAT"

  $reuse = Find-Master $Plan.MastersDir @("pp_flat_stacked.fit","masterflat.fit")
  if (-not $reuse) { $s.Copied = FlattenToCount $s.SrcDir $s.WorkDir }

  # Flat은 공통 master 개념 없음(세션별 생성/재사용만)
  Resolve-FromSessionOrCommon -Plan $Plan -State $s `
    -ReuseNames @("pp_flat_stacked.fit","masterflat.fit") `
    -CommonMasterPath $null `
    -CreateMasterName "pp_flat_stacked.fit"

  $Plan.Cal[$T_FLAT] = $s
}

function Resolve-Light {
  param([pscustomobject]$Plan)

  $s = New-CalState $T_LIGHT
  $s.SrcDir  = Join-Path $Plan.SPath "Light"
  $s.WorkDir = Join-Path $Plan.WorkRoot "LIGHT"

  # Light는 항상 프레임 필요(masters 재사용 개념 없음)
  $s.Copied = FlattenToCount $s.SrcDir $s.WorkDir
  if ($s.Copied -gt 0) { $s.Status = "복사" } else { $s.Status = "없음" }

  $Plan.Cal[$T_LIGHT] = $s
}

# =========================
# SSF 빌드
# =========================
function Build-SessionSsf {
  param([pscustomobject]$Plan)

  $B  = $Plan.Cal[$T_BIAS]
  $D  = $Plan.Cal[$T_DARK]
  $DF = $Plan.Cal[$T_DARKFLAT]
  $F  = $Plan.Cal[$T_FLAT]
  $L  = $Plan.Cal[$T_LIGHT]

  $ssf = @"
requires 1.2

"@

  # --- Bias master 생성 (필요시) ---
  if ($B.Create) {
    $ssf += @"
cd $(ToSirilQ $B.WorkDir)
convert bias
stack bias_ med -out=$(RelPath $B.WorkDir $B.Master)

"@
  }

# --- Dark master 생성 (필요시) : bias가 있으면 제거 후 stack --- bias는 제거하지 않음
  <# Gemini답변.
  우리가 찍는 Dark 프레임 안에는 이미 Bias(전기적 노이즈) 성분이 포함되어 있기 때문입니다.
  Dark = (열 노이즈) + (Bias 노이즈)
  Light = (별빛) + (열 노이즈) + (Bias 노이즈)
  따라서 Light에서 Dark를 그냥 빼버리면, 열 노이즈와 Bias가 한꺼번에 세트로 제거됩니다.
  $$Light - Dark = (별빛 + \text{열} + \text{Bias}) - (\text{열} + \text{Bias}) = \text{별빛}
  $$여기서 Dark를 만들기 전에 Bias를 미리 빼버린다면, 나중에 Light에서 뺄 때 Bias 성분이 남게 되어 별도로 또 Bias를 빼줘야 하는 번거로움이 생깁니다.
  #>
  if ($D.Create) {
    $ssf += @"
cd $(ToSirilQ $D.WorkDir)
convert dark

stack dark_ med -nonorm -out=$(RelPath $D.WorkDir $D.Master)

"@
	  
<#
    $darkBiasArg = ""
    if ($B.Master) { $darkBiasArg = " -bias=" + (RelPath $D.WorkDir $B.Master) }

    $ssf += @"
cd $(ToSirilQ $D.WorkDir)
convert dark

"@

    if ($darkBiasArg) {
      $ssf += @"
calibrate dark_$darkBiasArg -prefix=pp_
stack pp_dark_ med -nonorm -out=$(RelPath $D.WorkDir $D.Master)

"@
    } else {
      $ssf += @"
stack dark_ med -nonorm -out=$(RelPath $D.WorkDir $D.Master)

"@
    }
#>
  }

  # --- DarkFlat master 생성 (필요시) : bias가 있으면 제거 후 stack ---
  <# Gemini답변
  결론부터 말씀드리면, Darkflat을 만들 때도 Bias를 따로 뺄 필요가 없습니다.
  그 이유는 앞서 설명해 드린 Master Dark의 원리와 완전히 동일하기 때문입니다. Darkflat 자체가 이미 **'특정 노출 시간 동안 발생한 열 잡음 + Bias'**를 모두 머금고 있는 상태입니다.
  #>
  if ($DF.Create) {
    $ssf += @"
cd $(ToSirilQ $DF.WorkDir)
convert darkflat

stack darkflat_ med -nonorm -out=$(RelPath $DF.WorkDir $DF.Master)

"@

<#
    $dfBiasArg = ""
    if ($B.Master) { $dfBiasArg = " -bias=" + (RelPath $DF.WorkDir $B.Master) }

    $ssf += @"
cd $(ToSirilQ $DF.WorkDir)
convert darkflat

"@

    if ($dfBiasArg) {
      $ssf += @"
calibrate darkflat_$dfBiasArg -prefix=pp_
stack pp_darkflat_ med -nonorm -out=$(RelPath $DF.WorkDir $DF.Master)

"@
    } else {
      $ssf += @"
stack darkflat_ med -nonorm -out=$(RelPath $DF.WorkDir $DF.Master)

"@
    }
#>
  }

  # --- Flat master 생성 (필요시) ---
  <#
  Flat보정할 때 Light용 Dark를 써도 돼? 노출시간은 달라.
  Gemini said
  결론부터 말씀드리면, 절대 안 됩니다.
  천체사진 보정에서 가장 중요한 대원칙은 **"노출 시간(Exposure Time)이 다르면 Dark는 서로 호환되지 않는다"**는 것입니다. 
  #>
  if ($F.Create) {
    # Flat 보정 우선순위: 세션 DarkFlat > 세션 Dark > 공통 Dark --> 폐기
    # Flat 보정 우선순위: 세션 DarkFlat > 공통 Bias
	<#
    $FlatDarkArg = ""
    if ($DF.Master) { $FlatDarkArg = " -dark=" + (RelPath $F.WorkDir $DF.Master) }
    elseif ($D.Master) { $FlatDarkArg = " -dark=" + (RelPath $F.WorkDir $D.Master) }

    $FlatBiasArg = ""
    if ($B.Master) { $FlatBiasArg = " -bias=" + (RelPath $F.WorkDir $B.Master) }
	#>

    $FlatDarkArg = ""
    $FlatBiasArg = ""
    if ($DF.Master) { $FlatDarkArg = " -dark=" + (RelPath $F.WorkDir $DF.Master) } # DarkFlat을 사용하거나,
    elseif ($B.Master) { $FlatBiasArg = " -bias=" + (RelPath $F.WorkDir $B.Master) } # Bias를 사용한다
	
    $ssf += @"
cd $(ToSirilQ $F.WorkDir)
convert flat
calibrate flat_$FlatDarkArg$FlatBiasArg -cfa -prefix=cbflat_
stack cbflat_flat_ med -norm=mul -out=$(RelPath $F.WorkDir $F.Master)

"@
  }

  # --- Light calibrate + register ---
  # Light 보정 우선순위: 세션 Dark > 공통 Dark > 공통 Bias
  $calDark = ""
  $calBias = ""
  if ($D.Master) { $calDark = " -dark=" + (RelPath $L.WorkDir $D.Master) }
  elseif ($B.Master) { $calBias = " -bias=" + (RelPath $F.WorkDir $B.Master) } # Bias를 사용한다

  $calFlat = ""
  if ($F.Master) { $calFlat = " -flat=" + (RelPath $L.WorkDir $F.Master) }

  $ssf += @"
cd $(ToSirilQ $L.WorkDir)
convert light
calibrate light_$calDark$calBias$calFlat -cc=dark -cfa -equalize_cfa -debayer -prefix=pp_

# 별 적은 필드(B33) 등록 완화
#setfindstar -relax=on

register pp_light_

"@

  return $ssf
}

# =========================
# 세션 _WORK 정리(세션 실행 전)
# =========================
function Reset-Work([pscustomobject]$Plan) {
  if (Test-Path $Plan.WorkRoot) {
    Remove-Item $Plan.WorkRoot -Recurse -Force -ErrorAction SilentlyContinue
  }
  New-Item -ItemType Directory -Force -Path $Plan.WorkRoot | Out-Null
}

# =========================
# 메인
# =========================

# _ALL_REGISTERED 비우기
Get-ChildItem -Path $AllRegistered -File -ErrorAction SilentlyContinue |
  Remove-Item -Force -ErrorAction SilentlyContinue

if (-not (Test-Path $SessionRoot)) { throw "Session 폴더가 없습니다: $SessionRoot" }

# 1) 공통 Bias -> 공통 Dark(공통 Bias 반영)
$CreatedSsf = [System.Collections.Generic.List[string]]::new()

$CommonBiasFile = Ensure-CommonMaster -Kind "bias" -CommonDir $CommonBiasDir -HomeDir $HomeDir -SirilCli $SirilCli -CommonBiasForDark $null
if ($CommonBiasFile) { Write-Host "▶ 공통 Bias: $(Split-Path $CommonBiasFile -Leaf) 사용" }

$CommonDarkFile = Ensure-CommonMaster -Kind "dark" -CommonDir $CommonDarkDir -HomeDir $HomeDir -SirilCli $SirilCli -CommonBiasForDark $CommonBiasFile
if ($CommonDarkFile) { Write-Host "▶ 공통 Dark: $(Split-Path $CommonDarkFile -Leaf) 사용" }

# 2) 세션 루프
$report = New-Object System.Collections.Generic.List[string]
$sessionDirs = @(Get-ChildItem $SessionRoot -Directory | Sort-Object Name)
if ($sessionDirs.Count -eq 0) { throw "Session 하위에 세션 폴더가 없습니다." }

foreach ($sess in $sessionDirs) {
  $SPath = $sess.FullName
  $SName = $sess.Name

  $Plan = New-Plan -SPath $SPath -SName $SName -HomeDir $HomeDir -CommonBias $CommonBiasFile -CommonDark $CommonDarkFile
  New-Item -ItemType Directory -Force -Path $Plan.MastersDir | Out-Null

  # (중요) 세션 실행 전 _WORK 초기화
  Reset-Work $Plan

  # 타입별 Resolve (복사/재사용/공통/생성 계획 확정)
  Resolve-Bias     $Plan
  Resolve-Dark     $Plan
  Resolve-DarkFlat $Plan
  Resolve-Flat     $Plan
  Resolve-Light    $Plan

  $L = $Plan.Cal[$T_LIGHT]
  if ($L.Copied -le 0) {
    Write-Warning "세션 '$SName' → Light 없음, 건너뜀"
    continue
  }

  $B  = $Plan.Cal[$T_BIAS]
  $D  = $Plan.Cal[$T_DARK]
  $DF = $Plan.Cal[$T_DARKFLAT]
  $F  = $Plan.Cal[$T_FLAT]

  Write-Host ("▶ 세션: {0}  (DarkFlat:{1}, Dark:{2}, Flat:{3}, Bias:{4})" -f `
    $SName, $DF.Status, $D.Status, $F.Status, $B.Status)

  # SSF 생성/실행
  $ssf = Build-SessionSsf $Plan
  $ssfPath = Join-Path $SPath "___session_$SName.ssf"
  Write-NoBom $ssfPath $ssf
  $CreatedSsf.Add($ssfPath)

  $sessionLog = Join-Path $SPath "siril_session.log"
  $null = Invoke-Siril $SirilCli $ssfPath $sessionLog

  # 등록 결과 수집 (r_*)
  $registered = @(
    Get-ChildItem -Path $L.WorkDir -File -ErrorAction SilentlyContinue |
      Where-Object { $_.Name -like "r_*" -and ($_.Extension -in ".fit",".fits",".fts") }
  )

  if (-not $registered -or $registered.Count -eq 0) {
    Write-Warning "세션 '$SName' → 등록 결과(r_*.fit)가 없습니다. siril_session.log 확인 필요"
    $report.Add("$($SName): 0 registered frames")
  } else {
    foreach ($f in $registered) {
      $dst = Join-Path $AllRegistered ("{0}_{1}" -f $SName, $f.Name)
      Move-Item $f.FullName $dst -Force
    }
    $report.Add("$($SName): $($registered.Count) registered frames")
  }

  # 세션별 _WORK 정리(원하면 유지 가능)
  # Remove-Item $Plan.WorkRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# 3) 최종 스택
$allCount = (Get-ChildItem -Path $AllRegistered -File -ErrorAction SilentlyContinue |
             Where-Object { $_.Extension -in ".fit",".fits",".fts" } | Measure-Object).Count
if ($allCount -eq 0) { throw "_ALL_REGISTERED에 등록 프레임이 없습니다. 세션 처리/등록 단계 확인 필요." }

Write-Host ("▶ 최종: 등록 프레임 {0}개 → 1회 스택" -f $allCount)

# === 최종 스택(등록 프레임 전체 1회) ===
$finalSsf = @"
requires 1.2
cd $(ToSirilQ $AllRegistered)

# 등록 프레임 전체 스택
convert all
register all_
stack r_all_ rej 3 3 -norm=addscale -output_norm -rgb_equal -32b -out=stack_all

########################################
# 일반 처리 (stretch + 배경제거 + asinh)
########################################
load stack_all
save $(RelPath $AllRegistered $StackPath)

# 1) Histogram Auto Stretch
autostretch -linked

# 2) Background Extraction
subsky -rbf -samples=20 -smooth=0.5 -tolerance=2.0

# 3) Asinh Transform
asinh 1 0.2 -clipmode=globalrescale

# 4) 결과 저장
save stack_all_processed
savejpg stack_all_final 95
save $(RelPath $AllRegistered (Join-Path $HomeDir "stack_all_processed"))
savejpg $(RelPath $AllRegistered (Join-Path $HomeDir "stack_all_final")) 95

########################################
# HOO 처리 (R=Ha, G/B=OIII)
########################################
# 스택 원본(선형)을 다시 로드
load stack_all

# 1) 채널 분리 (stack_R/G/B 생성)
split stack_R stack_G stack_B

# 2) OIII 합성: G 비중↑, B 비중↓
pm '`$stack_G$ * 0.75 + `$stack_B$ * 0.25' -nosum
save OIII_lin

# 3) OIII 부스트 (2.0~3.0 범위에서 조절)
pm '`$OIII_lin$ * 2.5' -nosum
save OIII_boost

# 4) HOO 합성: R=stack_R, G=OIII, B=OIII
rgbcomp stack_R OIII_boost OIII_boost
load composed_rgb

# 5) 오토 스트레치
autostretch -linked

# 6) Background Extraction
subsky -rbf -samples=20 -smooth=0.5 -tolerance=2.0

# 7) Asinh
asinh 1 0.2 -clipmode=globalrescale

# 8) 저장
save stack_all_processed_HOO
savejpg stack_all_final_HOO 95
save $(RelPath $AllRegistered (Join-Path $HomeDir "stack_all_processed_HOO"))
savejpg $(RelPath $AllRegistered (Join-Path $HomeDir "stack_all_final_HOO")) 95

"@

$finalPath = Join-Path $HomeDir "___final_stack.ssf"
Write-NoBom $finalPath $finalSsf
$CreatedSsf.Add($finalPath)

Write-Host "▶ 최종 스택 실행..."
$finalLog = Join-Path $HomeDir "final_siril.log"
$null = Invoke-Siril $SirilCli $finalPath $finalLog
Write-Host "  · 최종 스택 로그: $finalLog"

# 4) 리포트 저장
$report.Insert(0, ("총 등록 프레임: {0}개" -f $allCount))
$reportPath = Join-Path $HomeDir "session_stack_report.txt"
[System.IO.File]::WriteAllLines($reportPath, $report)

# 5) 정리
Write-Host "▶ 중간 산출물 정리..."
# _WORK 폴더 삭제
Get-ChildItem -Path $SessionRoot -Directory -ErrorAction SilentlyContinue |
  ForEach-Object {
    $workDir = Join-Path $_.FullName '_WORK'
    if (Test-Path $workDir) {
      Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
    }
  }

# 홈 아래 중간 산출물 삭제
Get-ChildItem -Path (Join-Path $HomeDir '*') -File -Recurse -ErrorAction SilentlyContinue |
  Where-Object { $_.Name -like '*.ssf' -or $_.Name -like '*.seq' } |
  Remove-Item -Force

# 생성한 SSF 자동 삭제
if (-not $KeepSsf) {
  foreach ($p in $CreatedSsf) {
    if (Test-Path $p) { Remove-Item $p -Force -ErrorAction SilentlyContinue }
  }
}

Write-Host "`n? 완료! 최종 파일: $(Join-Path $HomeDir 'stack_all.fit')"
Write-Host "?? 세션별 프레임 수 요약: $reportPath"