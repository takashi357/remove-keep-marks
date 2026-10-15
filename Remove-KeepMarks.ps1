<#
.SYNOPSIS
  docx の見出しなどに付く「次の段落と分離しない」「段落を分割しない」「段落前で改ページ」を外す。

.DESCRIPTION
  これらの設定が付いた段落には、編集記号を表示した画面で行頭に黒い点が出る。
  点は印刷・PDF には出ない。消す必要があるのは、相手が編集記号を表示したまま
  docx を開く場合だけ。設定を外すと見出しがページ最下部に取り残されることがあるため、
  外したあとに見出しと次の段落が別ページに分かれていないかを確認して一覧にする。

  Word を画面に出さずに別途起動して処理する。元のファイルは上書きしない。

.PARAMETER Path
  処理する docx のパス。

.PARAMETER OutPath
  保存先。省略時は元ファイルと同じフォルダに「<元の名前>_nomarks.docx」で保存する。

.OUTPUTS
  終了コード 0：問題なし。1：失敗（途中で止まった）。2：保存したが、取り残された見出しがある。

.EXAMPLE
  .\Remove-KeepMarks.ps1 -Path .\resume.docx
#>
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [string]$OutPath
)

$ErrorActionPreference = 'Stop'

# 失敗したとき（ファイルが見つからない、Word を起動できない、開けない、保存できないなど）は終了コード 1
trap {
  [Console]::Error.WriteLine("エラー: $($_.Exception.Message)")
  exit 1
}

$src = (Resolve-Path -LiteralPath $Path).Path
if (-not $OutPath) {
  $dir  = Split-Path -Parent $src
  $name = [IO.Path]::GetFileNameWithoutExtension($src)
  $OutPath = Join-Path $dir ($name + '_nomarks.docx')
}
$OutPath = [IO.Path]::GetFullPath($OutPath)
if ($OutPath -eq $src) { throw '保存先が元ファイルと同じです。別の -OutPath を指定してください。' }

# Word の定数
$wdStyleTypeParagraph = 1
$wdStyleTypeLinked    = 6
$wdOutlineLevelBody   = 10
$wdActiveEndPageNumber = 3
$wdStatisticPages     = 2
$wdFormatDocumentDefault = 16

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $doc = $word.Documents.Open($src, $false, $true)   # 読み取り専用で開き、別名で保存する

  # 1. スタイル側の設定を外す（見出し 1 などは初期設定で付いている）
  foreach ($st in $doc.Styles) {
    try {
      if ($st.Type -ne $wdStyleTypeParagraph -and $st.Type -ne $wdStyleTypeLinked) { continue }
      $pf = $st.ParagraphFormat
      if ($pf.KeepWithNext -or $pf.KeepTogether -or $pf.PageBreakBefore) {
        $pf.KeepWithNext = 0; $pf.KeepTogether = 0; $pf.PageBreakBefore = 0
        Write-Host "スタイルから外した: $($st.NameLocal)"
      }
    } catch { }   # 組み込みスタイルの一部は読み取れないことがある
  }

  # 2. 段落ごとに付けた設定を外す
  $n = 0
  foreach ($p in $doc.Paragraphs) {
    if ($p.KeepWithNext -or $p.KeepTogether -or $p.PageBreakBefore) {
      $p.KeepWithNext = 0; $p.KeepTogether = 0; $p.PageBreakBefore = 0
      $n++
    }
  }
  Write-Host "段落から外した: $n 件"

  # 3. 見出しが次の段落と別ページに分かれていないか確認する
  $doc.Repaginate()
  $orphans = @()
  foreach ($p in $doc.Paragraphs) {
    if ($p.OutlineLevel -eq $wdOutlineLevelBody) { continue }
    $next = $p.Next()
    if (-not $next) { continue }
    $pg  = $doc.Range($p.Range.Start, $p.Range.Start).Information($wdActiveEndPageNumber)
    $npg = $doc.Range($next.Range.Start, $next.Range.Start).Information($wdActiveEndPageNumber)
    if ($pg -ne $npg) {
      $text = ($p.Range.Text -replace "`r", '').Trim()
      $orphans += [pscustomobject]@{ Page = $pg; Heading = $text.Substring(0, [Math]::Min(40, $text.Length)) }
    }
  }
  Write-Host "ページ数: $($doc.ComputeStatistics($wdStatisticPages))"

  $doc.SaveAs2($OutPath, $wdFormatDocumentDefault)
  $doc.Close([ref]0)
  Write-Host "保存した: $OutPath"

  if ($orphans.Count -gt 0) {
    Write-Warning 'ページ最下部に取り残された見出しがあります。Word で開いて直してください。'
    $orphans | Format-Table -AutoSize | Out-String | Write-Host
    $exitCode = 2
  } else {
    Write-Host 'ページ最下部に取り残された見出しはありません。'
    $exitCode = 0
  }
}
finally {
  $word.Quit([ref]0)
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
exit $exitCode
