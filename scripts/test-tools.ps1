# WP3 tool regression harness (plan §6.4): drives WordToolExecutor against a
# real Word instance with a generated test document (empty doc / Chinese text /
# headings / list / table). Usage: npm run test:tools
param(
    [string]$Dll = ""
)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot

foreach ($candidate in @($Dll, (Join-Path $root "apps\com-host\bin\x64h\OfficeAgent.ComHost.dll"), (Join-Path $root "apps\com-host\bin\x64\OfficeAgent.ComHost.dll"))) {
    if ($candidate -and (Test-Path $candidate)) { $Dll = $candidate; break }
}
if (-not $Dll -or -not (Test-Path $Dll)) { Write-Error "OfficeAgent.ComHost.dll not found"; exit 1 }
Write-Host "[test-tools] DLL: $Dll" -ForegroundColor Cyan

$asm = [System.Reflection.Assembly]::LoadFrom($Dll)
$null = [System.Reflection.Assembly]::LoadFrom((Join-Path (Split-Path $Dll) "Newtonsoft.Json.dll"))
$jobjectType = [Newtonsoft.Json.Linq.JObject]

$execType = $asm.GetType("OfficeAgent.ComHost.WordToolExecutor")
$flags = [System.Reflection.BindingFlags]"NonPublic,Static"
$execMethod = $execType.GetMethod("Execute", $flags)
$appField = $execType.GetField("App", $flags)

$script:passed = 0
$script:failed = 0

function Invoke-Tool($method, $payloadJson) {
    if ($method -eq "executeAction") { $payloadJson = '{"action":' + $payloadJson + '}' }
    $payload = [Newtonsoft.Json.Linq.JObject]::Parse($payloadJson)
    $r = $execMethod.Invoke($null, @($method, $payload))
    $rt = $r.GetType()
    $data = $rt.GetField("Data").GetValue($r)
    $dataJson = if ($data) { $data | ConvertTo-Json -Depth 10 -Compress } else { "" }
    return [pscustomobject]@{
        Success = [bool]$rt.GetField("Success").GetValue($r)
        Result  = [string]$rt.GetField("Result").GetValue($r)
        Error   = [string]$rt.GetField("Error").GetValue($r)
        Data    = $dataJson
    }
}

function Check($name, $result, $expect) {
    $text = $result.Result
    if (-not $text) { $text = $result.Data }
    $ok = $result.Success -and ($text -match $expect)
    if ($ok) {
        Write-Host "  [PASS] $name" -ForegroundColor Green
        $script:passed++
    } else {
        Write-Host "  [FAIL] $name" -ForegroundColor Red
        Write-Host "         success=$($result.Success) error=$($result.Error)" -ForegroundColor Yellow
        Write-Host "         result=$($result.Result.Substring(0, [Math]::Min(160, $result.Result.Length)))" -ForegroundColor Yellow
        $script:failed++
    }
}

$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Add()
    $appField.SetValue($null, $word)

    Write-Host "--- empty document ---" -ForegroundColor Cyan
    Check "get_document_stats (empty doc)" (Invoke-Tool "executeAction" '{"action":"get_document_stats","params":{}}') '"totalParagraphs"'

    Write-Host "--- content + headings (Chinese) ---" -ForegroundColor Cyan
    Check "insert_at_start (multi, heading1)" (Invoke-Tool "executeAction" '{"action":"insert_at_start","params":{"content":"项目概览\\n正文第一段","format":"heading1"}}') '已插入到文档开头'
    Check "insert_at_end (multi-line)" (Invoke-Tool "executeAction" '{"action":"insert_at_end","params":{"content":"结尾段落一\n结尾段落二"}}') '已追加到文档末尾'
    Check "read_document paragraph_range" (Invoke-Tool "executeAction" '{"action":"read_document","params":{"mode":"paragraph_range","paragraph_index":0,"count":5}}') '读取段落 0 起共'
    Check "get_document_stats (has heading)" (Invoke-Tool "executeAction" '{"action":"get_document_stats","params":{}}') '"headingCount"'
    Check "read_document heading_context" (Invoke-Tool "executeAction" '{"action":"read_document","params":{"mode":"heading_context","heading_text":"项目概览"}}') '及其子内容'

    Write-Host "--- edit tools ---" -ForegroundColor Cyan
    Check "insert_after_heading" (Invoke-Tool "executeAction" '{"action":"insert_after_heading","params":{"heading_text":"项目概览","content":"插入到标题后","format":"normal"}}') '已插入内容|后插入'
    Check "insert_after_paragraph" (Invoke-Tool "executeAction" '{"action":"insert_after_paragraph","params":{"paragraph_index":1,"content":"段后插入"}}') '已在段落 1 后插入'
    Check "replace_paragraph" (Invoke-Tool "executeAction" '{"action":"replace_paragraph","params":{"paragraph_index":0,"content":"新的标题"}}') '已替换段落 0'
    Check "set_paragraph_style" (Invoke-Tool "executeAction" '{"action":"set_paragraph_style","params":{"paragraph_index":1,"format":"heading2"}}') '已将段落 1 设置为'
    Check "get_paragraph_format" (Invoke-Tool "executeAction" '{"action":"get_paragraph_format","params":{"paragraph_index":0}}') '"paragraphIndex"'
    Check "merge_paragraphs" (Invoke-Tool "executeAction" '{"action":"merge_paragraphs","params":{"first_paragraph_index":1,"second_paragraph_index":2,"separator":" | "}}') '已合并段落'
    Check "find_and_replace" (Invoke-Tool "executeAction" '{"action":"find_and_replace","params":{"find_text":"结尾段落二","replace_text":"结尾段落二（改）"}}') '已替换'
    Check "find_and_replace_v2 (replace_all)" (Invoke-Tool "executeAction" '{"action":"find_and_replace_v2","params":{"find_text":"段落","replace_text":"节","replace_all":true,"match_case":false}}') '已替换'
    Check "apply_rich_format" (Invoke-Tool "executeAction" '{"action":"apply_rich_format","params":{"target_mode":"paragraph_index","paragraph_index":0,"font":{"bold":true,"size":14,"color":"#FF0000"}}}') '已应用富文本格式'
    Check "delete_paragraph" (Invoke-Tool "executeAction" '{"action":"delete_paragraph","params":{"paragraph_index":1}}') '已删除段落 1'

    Write-Host "--- tables ---" -ForegroundColor Cyan
    Check "insert_table (at_end)" (Invoke-Tool "executeAction" '{"action":"insert_table","params":{"location":"at_end","headers":["名称","数量"],"rows":[["苹果","3"],["香蕉","5"]]}}') '表格'
    Check "get_document_tables" (Invoke-Tool "executeAction" '{"action":"get_document_tables","params":{}}') '张表格'
    Check "read_table" (Invoke-Tool "executeAction" '{"action":"read_table","params":{"table_index":0}}') '苹果'
    Check "delete_table" (Invoke-Tool "executeAction" '{"action":"delete_table","params":{"table_index":0}}') '已删除表格 0'

    Write-Host "--- selection / context / misc ---" -ForegroundColor Cyan
    $word.Selection.HomeKey(6) | Out-Null  # wdStory=6 → go to start
    $word.Selection.EndKey(5) | Out-Null  # wdLine=5? select first line-ish
    Check "get_selection_info" (Invoke-Tool "executeAction" '{"action":"get_selection_info","params":{}}') '"isCursorOnly"'
    Check "getStructuredContext" (Invoke-Tool "getStructuredContext" '{}') 'totalParagraphs'
    Check "getWordContext" (Invoke-Tool "getWordContext" '{}') 'documentContext'
    Check "insert_at_cursor" (Invoke-Tool "executeAction" '{"action":"insert_at_cursor","params":{"content":"光标处内容"}}') '已在光标处插入内容'
    Check "undo_last_action" (Invoke-Tool "executeAction" '{"action":"undo_last_action","params":{"count":1}}') '已撤销最近'
    Check "reply_only" (Invoke-Tool "executeAction" '{"action":"reply_only","params":{}}') '仅回复文本'
    Check "task_complete" (Invoke-Tool "executeAction" '{"action":"task_complete","params":{"summary":"完成"}}') '完成'

    Write-Host "--- perf: large-document replace_all (Word drag-down regression) ---" -ForegroundColor Cyan
    $lines = 1..400 | ForEach-Object { "这是第 $_ 段测试文本，包含关键词苹果和香蕉，用于批量替换性能测试。" }
    $insertJson = @{ action = "insert_at_end"; params = @{ content = ($lines -join "\n") } } | ConvertTo-Json -Compress -Depth 5
    Check "perf: insert 400 paragraphs" (Invoke-Tool "executeAction" $insertJson) '已追加到文档末尾'
    $swatch = [System.Diagnostics.Stopwatch]::StartNew()
    $perf = Invoke-Tool "executeAction" '{"action":"find_and_replace_v2","params":{"find_text":"苹果","replace_text":"橙子","replace_all":true,"match_case":false}}'
    $swatch.Stop()
    Check "perf: replace_all x400 result" $perf '已替换 400 处匹配'
    $secs = [Math]::Round($swatch.Elapsed.TotalSeconds, 1)
    if ($secs -le 20) {
        Write-Host "  [PASS] perf: replace_all x400 elapsed ${secs}s" -ForegroundColor Green
        $script:passed++
    } else {
        Write-Host "  [FAIL] perf: replace_all x400 took ${secs}s (Word drag-down regression)" -ForegroundColor Red
        $script:failed++
    }

    Write-Host "--- simple insert + mermaid ---" -ForegroundColor Cyan
    Check "simpleInsert append_end" (Invoke-Tool "simpleInsert" '{"mode":"append_end","text":"简单追加"}') '已追加到文档末尾'
    Check "simpleInsert insert_start" (Invoke-Tool "simpleInsert" '{"mode":"insert_start","text":"简单开头"}') '已插入到文档开头'
    Check "simpleInsert replace" (Invoke-Tool "simpleInsert" '{"mode":"replace","text":"简单替换"}') '已替换当前选区'
    $png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
    Check "insert_mermaid_image (at_end)" (Invoke-Tool "executeAction" ('{"action":"insert_mermaid_image","params":{"mermaid_code":"graph TD; A-->B;","location":"at_end","image_data_url":"data:image/png;base64,' + $png + '"}}')) '插入流程图'
}
finally {
    if ($word) {
        try { $word.Documents | ForEach-Object { $_.Close(0) } } catch { }
        try { $word.Quit() } catch { }
    }
}

Write-Host ""
if ($script:failed -gt 0) {
    Write-Host "TOOL REGRESSION: $script:passed passed, $script:failed FAILED" -ForegroundColor Red
    exit 1
}
Write-Host "TOOL REGRESSION: $script:passed passed" -ForegroundColor Green
exit 0
