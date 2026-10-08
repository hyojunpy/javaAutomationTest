param(
    [string]$JdkHome = $env:JAVA_HOME,
    [string]$MavenHome,
    [string]$MavenRepository,
    [Parameter(Mandatory = $true)][string]$JavaFxJmods,
    [string]$OutputDirectory,
    [switch]$RunRegression
)
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$stamp = [guid]::NewGuid().ToString('N')
if (-not $JdkHome) { throw 'JDK 21 경로를 -JdkHome 또는 JAVA_HOME으로 지정하세요.' }
$java = Join-Path $JdkHome 'bin/java.exe'
$jpackage = Join-Path $JdkHome 'bin/jpackage.exe'
if (-not (Test-Path -LiteralPath $jpackage)) { throw "jpackage.exe가 없습니다: $jpackage" }
if (-not (Test-Path -LiteralPath (Join-Path $JavaFxJmods 'javafx.controls.jmod'))) { throw 'JavaFX 21.0.4 Windows x64 JMODs 경로를 확인하세요.' }
$version = & $java --version | Out-String
if ($LASTEXITCODE -ne 0 -or $version -notmatch '(?:openjdk|java)\s+21[.\s]') { throw "JDK 21이 필요합니다: $version" }
$maven = if ($MavenHome) { Join-Path $MavenHome 'bin/mvn.cmd' } else { (Get-Command mvn.cmd -ErrorAction Stop).Source }
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $projectRoot "portable/$stamp" }
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath (Join-Path $OutputDirectory 'Jungwha_Converter')) { throw '출력 폴더가 이미 있습니다. 새 경로를 지정하세요.' }
$previousJava = $env:JAVA_HOME
$previousPath = $env:PATH
Push-Location $projectRoot
try {
    $env:JAVA_HOME = $JdkHome
    $env:PATH = "$JdkHome\bin;$env:PATH"
    $buildArguments = if ($RunRegression) { @('-Pregression', 'verify') } else { @('package') }
    if ($MavenRepository) { $buildArguments = @("-Dmaven.repo.local=$MavenRepository") + $buildArguments }
    & $maven @buildArguments
    if ($LASTEXITCODE -ne 0) { throw 'Maven 빌드 또는 테스트 실패' }
    $packageInput = Join-Path $projectRoot "target/packaging/$stamp"
    New-Item -ItemType Directory -Path $packageInput -Force | Out-Null
    Copy-Item -LiteralPath 'target/menu-hwp-to-excel-1.0.0-all.jar' -Destination $packageInput
    $arguments = @(
        '--type', 'app-image', '--name', 'Jungwha_Converter', '--app-version', '1.0.0',
        '--dest', $OutputDirectory, '--input', $packageInput,
        '--main-jar', 'menu-hwp-to-excel-1.0.0-all.jar', '--main-class', 'org.example.FxApp',
        '--icon', (Join-Path $projectRoot 'icon.ico'),
        '--module-path', "$JdkHome\jmods;$JavaFxJmods",
        '--add-modules', 'java.se,jdk.unsupported,jdk.charsets,javafx.controls,javafx.graphics',
        '--java-options', '--add-modules=javafx.controls,javafx.graphics',
        '--java-options', '-Dfile.encoding=UTF-8'
    )
    & $jpackage @arguments
    if ($LASTEXITCODE -ne 0) { throw 'EXE 패키징 실패' }
    $app = Join-Path $OutputDirectory 'Jungwha_Converter'
    Copy-Item -LiteralPath (Join-Path $projectRoot 'input') -Destination (Join-Path $app 'input') -Recurse
    Copy-Item -LiteralPath (Join-Path $projectRoot 'docs/사용방법.txt') -Destination $app
    Compress-Archive -LiteralPath $app -DestinationPath (Join-Path $OutputDirectory 'Jungwha_Converter-Windows-x64.zip')
    Write-Host "완료: $app"
} finally {
    Pop-Location
    $env:JAVA_HOME = $previousJava
    $env:PATH = $previousPath
}
