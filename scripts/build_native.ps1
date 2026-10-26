param([switch]$Test, [switch]$Publish, [string]$Version = '0.1.0')
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Push-Location $root
try {
    $dotnet = (Get-Command dotnet.exe -ErrorAction Stop).Source
    & $dotnet restore Qingxu.Tests/Qingxu.Tests.csproj --configfile Qingxu/NuGet.Config
    if ($LASTEXITCODE) { throw '项目还原失败，请确认已安装 .NET 10 SDK。' }
    & $dotnet build Qingxu.Tests/Qingxu.Tests.csproj -c Release --no-restore
    if ($LASTEXITCODE) { throw '构建失败。' }

    if ($Test) {
        $book = Join-Path $root 'outputs\study-plan\计划.xlsx'
        if (-not (Test-Path -LiteralPath $book)) {
            throw '测试需要一个工作簿：请把 templates\计划模板.xlsx 复制为 outputs\study-plan\计划.xlsx'
        }
        $result = & $dotnet Qingxu.Tests/bin/Release/net10.0-windows/Qingxu.Tests.dll $book 2>&1
        if ($LASTEXITCODE) {
            $result | Select-Object -Last 3 | ForEach-Object { Write-Host $_ }
            throw '回归测试未通过。注意：空白模板不含断言所需的具体课程和任务；请改用一份已填写、符合模板约定的工作簿，或省略 -Test 直接 -Publish。'
        }
        $result | ForEach-Object { Write-Host $_ }
    }

    if ($Publish) {
        & $dotnet restore Qingxu/Qingxu.csproj -r win-x64 --configfile Qingxu/NuGet.Config --source https://api.nuget.org/v3/index.json -p:SelfContained=true
        if ($LASTEXITCODE) { throw '微软运行时包还原失败。' }
        & $dotnet publish Qingxu/Qingxu.csproj -c Release -r win-x64 --self-contained true --no-restore `
            -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true `
            -p:EnableCompressionInSingleFile=false -p:DebugType=none -o dist/native-win-x64
        if ($LASTEXITCODE) { throw '自包含发布失败。' }

        $name = 'Qingxu-Native-v' + $Version
        $stage = Join-Path $root ('dist\' + $name)
        if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Recurse -Force }
        [void][IO.Directory]::CreateDirectory($stage)
        Copy-Item -LiteralPath (Join-Path $root 'dist\native-win-x64\青序.exe') -Destination $stage
        Copy-Item -LiteralPath (Join-Path $root 'templates\计划模板.xlsx') -Destination $stage
        Copy-Item -LiteralPath (Join-Path $root 'README.md') -Destination $stage
        Copy-Item -LiteralPath (Join-Path $root '创建桌面快捷方式（原生版）.ps1') -Destination $stage

        $manifest = Get-ChildItem -LiteralPath $stage -Recurse -File |
            Where-Object { $_.Name -ne 'SHA256SUMS.txt' } |
            ForEach-Object { (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash + '  ' + $_.Name }
        [IO.File]::WriteAllLines((Join-Path $stage 'SHA256SUMS.txt'), $manifest, [Text.Encoding]::UTF8)

        $zip = Join-Path $root ('dist\' + $name + '-win-x64.zip')
        if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }
        Compress-Archive -LiteralPath $stage -DestinationPath $zip -CompressionLevel Optimal
        Write-Host ('发布包：' + $zip)
        Get-FileHash -LiteralPath $zip -Algorithm SHA256 | Format-List
    }
}
finally { Pop-Location }
