<#
.SYNOPSIS
  Remove-KeepMarks.ps1 を試すための docx を作る。個人情報は含まない。

.DESCRIPTION
  次の 2 つを 1 つの文書に入れる。
  1. 見出し 1 の段落に、直接「段落を分割しない」を設定する
     （段落ごとの設定を外す処理が動くかを確かめる）。
  2. 見出し 2 を、keep 系の設定を外したときにちょうど 1 ページ目の最下部に来る位置に置く
     （取り残された見出しを検出できるかを確かめる）。
     見出しスタイルの初期設定が残っている間は、Word が見出しを次のページへ送るので取り残されない。

.PARAMETER OutPath
  保存先。省略時はこのスクリプトと同じフォルダの example-keep-marks.docx。
#>
param(
  [string]$OutPath = (Join-Path $PSScriptRoot 'example-keep-marks.docx')
)

$ErrorActionPreference = 'Stop'
$OutPath = [IO.Path]::GetFullPath($OutPath)

$wdStyleHeading1 = -2
$wdStyleHeading2 = -3
$wdStyleNormal   = -1
$wdActiveEndPageNumber = 3
$wdFormatDocumentDefault = 16

function Get-StartPage($doc, $para) {
  $doc.Range($para.Range.Start, $para.Range.Start).Information($wdActiveEndPageNumber)
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $doc = $word.Documents.Add()
  $doc.RemovePersonalInformation = $true   # 作成者などを保存しない

  # 見出し 1 の「段落を分割しない」はスタイルから外し、段落に直接付ける。
  # スタイルと同じ値を段落に付けても、直接の設定として保存されないため。
  $doc.Styles.Item($wdStyleHeading1).ParagraphFormat.KeepTogether = 0

  $doc.Content.Text = "試験用の文書`r見出し：ページ最下部に来る見出し`rこの段落は見出しの直後の本文。"
  $title = $doc.Paragraphs.Item(1)
  $title.Style = $wdStyleHeading1
  $title.KeepTogether = -1
  $doc.Paragraphs.Item(2).Style = $wdStyleHeading2
  $doc.Paragraphs.Item(3).Style = $wdStyleNormal

  # 見出し 2 の keep 系を一時的に直接外し、見出しが 1 ページ目の最下部、
  # 本文が 2 ページ目の先頭になるまで、見出しの前に本文の行を足していく。
  $n = 0
  while ($true) {
    $count = $doc.Paragraphs.Count
    $h = $doc.Paragraphs.Item($count - 1)
    $h.KeepWithNext = 0; $h.KeepTogether = 0
    $doc.Repaginate()
    $hp = Get-StartPage $doc $h
    $bp = Get-StartPage $doc $doc.Paragraphs.Item($count)
    if ($hp -eq 1 -and $bp -eq 2) { break }
    if ($hp -ge 2) { throw "見出しを 1 ページ目の最下部に置けなかった（本文の行 $n 行）" }
    if ($n -ge 200) { throw '本文の行を 200 行足しても条件を満たさなかった' }
    $n++
    $doc.Range($h.Range.Start, $h.Range.Start).InsertBefore("本文の行 $n。`r")
    $doc.Paragraphs.Item($count - 1).Style = $wdStyleNormal   # 足した行（見出しの直前）
  }

  # 一時的に付けた直接の設定を戻し、スタイルの初期設定（次の段落と分離しない）を効かせる
  $h = $doc.Paragraphs.Item($doc.Paragraphs.Count - 1)
  $h.Range.ParagraphFormat.Reset()
  $doc.Repaginate()
  Write-Host "本文の行: $n 行、見出し 2 の開始ページ（初期設定のまま）: $(Get-StartPage $doc $h)"

  $doc.SaveAs2($OutPath, $wdFormatDocumentDefault)
  $doc.Close([ref]0)
  Write-Host "保存した: $OutPath"
}
finally {
  $word.Quit([ref]0)
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
