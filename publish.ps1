# Publishes every tool into one folder, so adding that folder to PATH adds them all. The folder is ToolsPublishDir in
# Directory.Build.props. A project is published if it has a FolderProfile, so a new tool needs nothing added here.
$ErrorActionPreference = "Stop"

$publishProfiles = Get-ChildItem "$PSScriptRoot/*/src/*/Properties/PublishProfiles/FolderProfile.pubxml"

try {
    foreach ($publishProfile in $publishProfiles) {
        $projectFolder = $publishProfile.Directory.Parent.Parent
        $project = Get-ChildItem $projectFolder.FullName -Filter *.csproj

        Write-Host "Publishing $($project.BaseName)"
        dotnet publish $project.FullName -c Release -p:PublishProfile=FolderProfile
        if ($LASTEXITCODE -ne 0) {
            exit $LASTEXITCODE
        }
    }
}
finally {
    # dotnet keeps its build servers (MSBuild's worker nodes and the compiler's server) running for the next build, and
    # Task Manager shows them as ".NET Host". Publishing is usually the last build for a while, so they're ended here,
    # whether it worked or not. A failed publish still exits with its own code.
    Write-Host "Ending the build servers"
    dotnet build-server shutdown | Out-Null
}
