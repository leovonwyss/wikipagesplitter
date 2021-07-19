$wikiBaseUrl = "https://wiki.migros.net"
if (-not $wikiCredentials) {
    $wikiCredentials = Get-Credential -Message "Bitte geben Sie Nutzername und Passwort für den Zugriff auf Confluence ein"
}

function Invoke-WikiRequest {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $relativePath,

        [Parameter()]
        $method,

        [Parameter()]
        [string]$jsonContent
    )
    
    if ($relativePath.StartsWith("https://")) {
        $finalUrl = $relativePath;
    } else {
        $finalUrl = "$($wikiBaseUrl.TrimEnd('/'))/$($relativePath.TrimStart('/'))";
    }

    if ($jsonContent) {
        return Invoke-WebRequest $finalUrl -UseBasicParsing -Credential $wikiCredentials -Authentication Basic -Method $method -Body ([System.Text.Encoding]::UTF8.GetBytes($jsonContent)) -ContentType "application/json" -Verbose #-proxy "http://webproxy.dc.migros.ch:9099"
    } else {
        return Invoke-WebRequest $finalUrl -UseBasicParsing -Credential $wikiCredentials -Authentication Basic #-proxy "http://webproxy.dc.migros.ch:9099"
    }
}

function Write-NewConfluencePage {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $title,
        [Parameter(Mandatory=$true)]
        [string]
        $spaceKey,
        [Parameter(Mandatory=$true)]
        [int]
        $parentPageId,
        [Parameter(Mandatory=$true)]
        [string]
        $content
    )
    
    $newPageObject = [PSCustomObject]@{
        type = "page";
        title = $title;
        ancestors = ([array]@([PSCustomObject]@{ id = $parentPageId }));
        space = [PSCustomObject]@{ key = $spaceKey };
        body = [PSCustomObject]@{ 
            storage = [PSCustomObject]@{ 
                value = $content; 
                representation = "storage" 
            }
        }
    };
    $jsonContent = [string](ConvertTo-Json $newPageObject);
    [System.IO.File]::WriteAllText("payload.json", $jsonContent)

    return Invoke-WikiRequest "/rest/api/content/" -Method POST -jsonContent $jsonContent
}

function Write-ExistingConfluencePage {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [int]
        $pageId,
        [Parameter(Mandatory=$true)]
        [int]
        $currentVersion,
        [Parameter(Mandatory=$true)]
        [string]
        $title,
        [Parameter(Mandatory=$true)]
        [string]
        $spaceKey,
        [Parameter(Mandatory=$true)]
        [string]
        $content
    )

    $newPageObject = [PSCustomObject]@{
        type = "page";
        title = $title;
        space = [PSCustomObject]@{ key = $spaceKey };
        body = [PSCustomObject]@{ 
            storage = [PSCustomObject]@{ 
                value = $content; 
                representation = "storage" 
            }
        };
        version = [PSCustomObject]@{ number = ($currentVersion + 1) };
    };
    $jsonContent = [string](ConvertTo-Json $newPageObject);
    [System.IO.File]::WriteAllText("payload.json", $jsonContent)

    return Invoke-WikiRequest "/rest/api/content/$($pageId)" -Method PUT -jsonContent $jsonContent
}

function WrapInHtmlContainer{
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $content
    )
    
    return @"
<!DOCTYPE html [
$(Get-Content "$PSScriptRoot\xhtml-lat1.ent" -Raw)
$(Get-Content "$PSScriptRoot\xhtml-special.ent" -Raw)
$(Get-Content "$PSScriptRoot\xhtml-symbol.ent" -Raw)
]>
<html xmlns:ac='atlassian-confluence' xmlns:ri='atlassian-confluence-ri'>
$content
</html>
"@    
}

$sourcePageUrl = "https://wiki.migros.net/pages/viewpage.action?pageId=441044511"

# get page id
$sourcePageRendered = [string](Invoke-WikiRequest $sourcePageUrl)
if ($sourcePageRendered -match '\<meta\sname="ajs-page-id"\scontent="(?<pageid>\d+)"') {
    $pageid = $matches.pageid
    Write-Output "Seite $pageid wird verarbeitet..."
} else {
    Write-Error "Die Seitennummer konnte nicht ausgelesen werden. Bitte Url prüfen."
}
if ($sourcePageRendered -match '\<meta\sname="ajs-page-version"\scontent="(?<pageversion>\d+)"') {
    $pageVersion = $matches.pageversion
} else {
    Write-Error "Die Seitenversion konnte nicht ausgelesen werden. Bitte Url prüfen."
}
if ($sourcePageRendered -match '\<meta\sname="ajs-page-title"\scontent="(?<pagetitle>[^"]+)"') {
    $pageTitle = $matches.pagetitle
} else {
    Write-Error "Der Space-Key konnte nicht ausgelesen werden. Bitte Url prüfen."
}
if ($sourcePageRendered -match '\<meta\sname="ajs-space-key"\scontent="(?<spacekey>\w+)"') {
    $spacekey = $matches.spacekey
} else {
    Write-Error "Der Space-Key konnte nicht ausgelesen werden. Bitte Url prüfen."
}

#get page source
$sourcePageData = Invoke-WikiRequest "/rest/api/content/$($pageid)?expand=body.storage"
$sourcePageBody = (ConvertFrom-Json $sourcePageData.Content).body.storage.value
[System.IO.File]::WriteAllText("sourcepage.xml", (WrapInHtmlContainer($sourcePageBody)))

$sourcePageBodyXml = [System.Xml.XmlDocument]::new()
$sourcePageBodyXml.LoadXml((Get-Content "sourcepage.xml" -Raw))

$childPageNumber = 1
$h1Nodes = $sourcePageBodyXml.DocumentElement.SelectNodes("//h1")
$h1Nodes | ForEach-Object {
    # Leeres Dokument erstellen als Ziel
    $newDocument = [System.Xml.XmlDocument]::new()
    $newDocument.LoadXml((WrapInHtmlContainer('')))

    $nodeToTransfer = $_
    $nodeToRemove = $nodeToTransfer.SelectSingleNode("*");
    while ($nodeToRemove) {
        $removed = $nodeToTransfer.RemoveChild($nodeToRemove);
        $nodeToRemove = $nodeToTransfer.SelectSingleNode("*");
    }
    $newPageTitle = "$($childPageNumber.ToString('00')). $($_.InnerText) ($pageid)"

    $nodeToTransfer = $_.NextSibling # die h1 an sich nicht übertragen
    while ($nodeToTransfer) {
        $thisParent = $nodeToTransfer.ParentNode
        $nextNode = $nodeToTransfer.NextSibling

        $removed = $nodeToTransfer.ParentNode.RemoveChild($nodeToTransfer);

        if ($removed.LocalName -match "^h(?<headingLevel>\d)$") {
            $removedContents = $removed.InnerXml;
            $removed = $newDocument.CreateElement("h" + (([int]$matches.headingLevel) - 1));
            $removed.InnerXml = $removedContents;
        }

        $added = $newDocument.DocumentElement.AppendChild($newDocument.ImportNode($removed, $true));

        if (($nextNode.LocalName -eq "h1") -or ($nextNode.ParentNode -ne $thisParent)) {
            $nodeToTransfer = $null
        } else {
            $nodeToTransfer = $nextNode
        }
    }
    
    Write-Output "-----"
    Write-Output "h1: $($_.InnerText)"
    Write-Output $newDocument.OuterXml

    Write-NewConfluencePage -spaceKey $spaceKey -parentPageId $pageid -title $newPageTitle -content $newDocument.DocumentElement.InnerXml

    $childPageNumber = $childPageNumber+1
}

$remainingContent = $sourcePageBodyXml.DocumentElement.InnerXml
Write-ExistingConfluencePage -pageId $pageid -spaceKey $spacekey -currentVersion $pageVersion -title $pageTitle -content $remainingContent
