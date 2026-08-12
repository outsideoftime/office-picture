param(
    [ValidatePattern('^\d+\.\d+\.\d+\.\d+$')]
    [string]$Version = '1.0.0.0'
)

$ErrorActionPreference = 'Stop'

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$wordProjectPath = Join-Path $repositoryRoot 'src\OfficePicture.WordAddIn\OfficePicture.WordAddIn.csproj'
[xml]$wordProject = Get-Content -LiteralPath $wordProjectPath

$keyNode = $wordProject.SelectSingleNode("//*[local-name()='ManifestKeyFile']")
$thumbprintNode = $wordProject.SelectSingleNode("//*[local-name()='ManifestCertificateThumbprint']")
if ($null -eq $keyNode -or $null -eq $thumbprintNode) {
    throw 'Create a ClickOnce signing certificate in the Word project before publishing.'
}

$keyPath = Join-Path (Split-Path -Parent $wordProjectPath) $keyNode.InnerText
if (-not (Test-Path -LiteralPath $keyPath)) {
    throw "Signing certificate not found: $keyPath"
}

$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path -LiteralPath $vswhere)) {
    throw 'Visual Studio Installer vswhere.exe was not found.'
}

$visualStudioPath = & $vswhere -latest -products '*' -requires 'Microsoft.Component.MSBuild' -property installationPath
if ([string]::IsNullOrWhiteSpace($visualStudioPath)) {
    throw 'Visual Studio MSBuild was not found.'
}

$msbuild = Join-Path $visualStudioPath 'MSBuild\Current\Bin\MSBuild.exe'
$publishRoot = Join-Path $repositoryRoot 'publish'
$projects = @(
    @{ Host = 'Word'; Project = 'src\OfficePicture.WordAddIn\OfficePicture.WordAddIn.csproj'; Product = 'OfficePicture Word' },
    @{ Host = 'PowerPoint'; Project = 'src\OfficePicture.PowerPointAddIn\OfficePicture.PowerPointAddIn.csproj'; Product = 'OfficePicture PowerPoint' },
    @{ Host = 'Excel'; Project = 'src\OfficePicture.ExcelAddIn\OfficePicture.ExcelAddIn.csproj'; Product = 'OfficePicture Excel' }
)

foreach ($item in $projects) {
    $projectPath = Join-Path $repositoryRoot $item.Project
    $publishDirectory = (Join-Path $publishRoot $item.Host) + '\'
    $arguments = @(
        $projectPath,
        '/t:Publish',
        '/p:Configuration=Release',
        '/p:Platform=AnyCPU',
        '/p:SignManifests=true',
        "/p:ManifestKeyFile=$keyPath",
        "/p:ManifestCertificateThumbprint=$($thumbprintNode.InnerText)",
        "/p:PublishDir=$publishDirectory",
        "/p:PublishUrl=$publishDirectory",
        "/p:ApplicationVersion=$Version",
        "/p:ProductName=$($item.Product)",
        '/p:PublisherName=OfficePicture',
        '/p:IsWebBootstrapper=false',
        '/p:BootstrapperEnabled=true',
        '/p:BootstrapperComponentsLocation=HomeSite',
        '/v:minimal'
    )

    Write-Host "Publishing $($item.Host) $Version ..."
    & $msbuild @arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$($item.Host) publish failed. MSBuild exit code: $LASTEXITCODE"
    }
}

Write-Host "Publish completed: $publishRoot"
