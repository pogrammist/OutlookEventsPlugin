# Скрипт для публикации и установки надстройки Outlook
param(
    [Parameter(Mandatory=$false)]
    [string]$Configuration = "Release"
)

# Пути
$projectPath = $PSScriptRoot
$publishPath = Join-Path $projectPath "publish"
$msbuildPath = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"

# Очистка предыдущей публикации
if (Test-Path $publishPath) {
    Remove-Item -Path $publishPath -Recurse -Force
}

# Публикация проекта
& $msbuildPath "$projectPath\OutlookEventsPlugin.csproj" /p:Configuration=$Configuration /p:PublishDir=$publishPath /t:Publish

# Создание установщика
$setupPath = Join-Path $publishPath "setup.exe"
$vstoPath = Join-Path $publishPath "OutlookEventsPlugin.vsto"

# Создание ярлыка для установки
$shortcutPath = Join-Path $publishPath "Install.lnk"
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($shortcutPath)
$Shortcut.TargetPath = $vstoPath
$Shortcut.Save()

Write-Host "Публикация завершена. Файлы находятся в папке: $publishPath"
Write-Host "Для установки запустите файл: $vstoPath" 
