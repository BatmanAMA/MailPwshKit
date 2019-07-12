param (
    $Configuration = "Debug",
    $ModuleName = (Get-Item .).BaseName
)
$BasePath = if ($PSScriptRoot){$PSScriptRoot} else {$PWD}
$ModuleInfo = Test-ModuleManifest -Path (Join-Path -Path $BasePath -ChildPath "$ModuleName.psd1")
$ModuleOutputPath = Join-Path $BasePath -ChildPath $ModuleName -AdditionalChildPath $ModuleInfo.Version
$pkgRepoName = "TMPREPO-$(New-Guid)"
task preClean -Before Build {
    remove $moduleOutputPath,bin,obj
}
task Build {
    dotnet.exe publish -c $configuration -o $ModuleOutputPath
    Copy-Item -Path $ModuleInfo.Path -Destination $ModuleOutputPath
    $ManifestData = @{
        Path = Join-Path -Path $ModuleOutputPath -ChildPath "$ModuleName.psd1"
        RootModule = 'mailpwshkit.dll'
        RequiredAssemblies = (Get-ChildItem $ModuleOutPutPath\*.dll).Name
        FileList = (Get-ChildItem $ModuleOutputPath -File -Recurse).FullName -replace [Regex]::Escape($ModuleOutputPath),'.'
    }
    Update-ModuleManifest @ManifestData
}
task Clean -after Build {
    remove bin,obj
}
task Compress -If ($configuration -eq 'Release' -and $ENV:CI) -after Build {
    Compress-Archive -Path (Join-Path $BasePath -ChildPath $ModuleName) -DestinationPath "$ModuleName.zip"
}

task Package Build,{
    $null = New-Item -Name "pkg" -ItemType Directory -Path $BasePath -ErrorAction SilentlyContinue
    $pkg = get-item pkg
    Register-PSRepository -Name $pkgRepoName -SourceLocation $pkg.FullName -PublishLocation $pkg.FullName -InstallationPolicy Trusted
    Publish-Module -Path "$BasePath\$ModuleName" -Repository $pkgRepoName
}

task Install Package,{
    Install-Module $ModuleName -Force
},PkgClean

task PkgClean -after Install {
    Unregister-PSRepository -Name $pkgRepoName -Force
    remove pkg
}

task . Build