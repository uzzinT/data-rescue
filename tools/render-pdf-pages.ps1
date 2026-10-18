# PDF의 원하는 쪽을 PNG 이미지로 뽑습니다 (Windows 내장 PDF 엔진 사용, 별도 설치 없음).
# 예: .\tools\render-pdf-pages.ps1 -Pdf "C:\...\통합과학2_지도서.pdf" -Pages 242,243,244 -OutDir "C:\...\pages" -Width 1300
# 지도서 쪽 번호와 PDF 쪽 번호가 같습니다 (비상교육 통합과학2 지도서 기준). 뽑은 이미지는 저작권 때문에 저장소에 올리지 마세요.
param([string]$Pdf, [int[]]$Pages, [string]$OutDir, [int]$Width = 1100)
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
$null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
function Await($op, [Type]$t) { $task = $asTaskGeneric.MakeGenericMethod($t).Invoke($null, @($op)); $task.Wait(-1) | Out-Null; $task.Result }
function AwaitAction($op) { $m = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' }; $t = $m[0].Invoke($null, @($op)); $t.Wait(-1) | Out-Null }
$file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($Pdf)) ([Windows.Storage.StorageFile])
$doc = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
"pages: $($doc.PageCount)"
foreach ($p in $Pages) {
  $page = $doc.GetPage($p - 1)
  $opts = New-Object Windows.Data.Pdf.PdfPageRenderOptions
  $opts.DestinationWidth = $Width
  $ms = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
  AwaitAction ($page.RenderToStreamAsync($ms, $opts))
  $net = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($ms.GetInputStreamAt(0))
  $out = Join-Path $OutDir ("p{0:D3}.png" -f $p)
  $fs = [IO.File]::Create($out); $net.CopyTo($fs); $fs.Close()
  $out
}
