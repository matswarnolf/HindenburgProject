function Add-Episode {
    param (
        [String]$Episode
    )
    $PodcastRoot = 'C:\Users\MatsWarnolf\OneDrive - Mats Warnolf AB\Documents\Office 365-podden\'
    If (!$Episode) {
        $Episode = Read-Host -Prompt 'What Episode would you Like to create'
    }
    $PodcastTemplateDirectory = "$PodcastRoot\Office365Podden.template\"
    $fileName = "$podcastTemplateDirectory\TemplateProject.nhsx";
    $xmlDoc = [System.Xml.XmlDocument](Get-Content $fileName)
    $EpisodePath = "$PodcastRoot" + "episoder\" 
    $EpisodeFilePath = "$EpisodePath\$episode Files"
    New-Item -Path $EpisodeFilePath -ItemType Directory -Force
    $xmlDoc.Session.AudioPool.Path = "$episode Files"
    $xmlDoc.Session.AudioPool.Location = "$PodcastRoot$Episode"
    $xmlDoc.Save("$EpisodePath\$Episode.nhsx")
    Copy-Item -Path "$PodcastTemplateDirectory\TemplateProject Files\*" -Exclude '*.nhsx' -Destination $EpisodeFilePath -Recurse
    $ExistingProjects = Get-ChildItem -Path "$PodcastRoot\Episoder" -Directory -Force -ErrorAction SilentlyContinue | Select-Object Name
    Write-host "Existing Episodes are:"
    $ExistingProjects | Format-List

}