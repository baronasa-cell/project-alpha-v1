# Ensure src folder exists
if (-not (Test-Path -Path src)) { New-Item -ItemType Directory -Path src | Out-Null }

# 1. app_js.html
$appContent = Get-Content app.js -Raw -Encoding UTF8
$appHtml = "<script>`n" + $appContent + "`n</script>"
Set-Content -Path src/app_js.html -Value $appHtml -Encoding UTF8

# 1.5. appsscript.json (マニフェストファイルをコピー)
if (Test-Path -Path appsscript.json) {
    Copy-Item -Path appsscript.json -Destination src/appsscript.json -Force
}

# 2. style_css.html (style.css の内容をそのまま使用)
$styleContent = Get-Content style.css -Raw -Encoding UTF8
$styleHtml = "<style>`n" + $styleContent + "`n</style>"
Set-Content -Path src/style_css.html -Value $styleHtml -Encoding UTF8

# 3. index.html (GAS用のinclude構文に置換)
$indexContent = Get-Content index.html -Raw -Encoding UTF8
$indexContent = $indexContent -replace '<link rel="stylesheet" href="style.css">', "<?!= include('style_css'); ?>"
$indexContent = $indexContent -replace '<script src="app.js"></script>', "<?!= include('app_js'); ?>"
Set-Content -Path src/index.html -Value $indexContent -Encoding UTF8

# 4. backend.gs (ルートのファイルをsrcにコピー)
$backendContent = Get-Content backend.gs -Raw -Encoding UTF8
Set-Content -Path src/backend.gs -Value $backendContent -Encoding UTF8
