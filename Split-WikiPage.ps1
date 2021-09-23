#
# Split-ConfluencePage
#

param (
    [Parameter(Mandatory=$true)]
    [string]
    $pageid
)

$ErrorActionPreference = 'Stop'

$wikiBaseUrl = "https://wiki.migros.net/"
if (-not $wikiCredentials) {
    $wikiCredentials = Get-Credential -Message "Bitte geben Sie Nutzername und Passwort für den Zugriff auf Confluence ein"
    $pair = "$($wikiCredentials.UserName):$($wikiCredentials.GetNetworkCredential().Password)"
    $encodedCreds = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($pair))
    $wikiBasicAuthValue = "Basic $encodedCreds"
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
        [string]$jsonContent,

        [Parameter()]
        [string]$outFile,

        [Parameter()]
        [string]$attachFile
    )
    
    if ($relativePath.StartsWith("https://")) {
        $finalUrl = $relativePath;
    } else {
        if (-not $wikiBaseUrl) {
            Write-Error "Wiki Base Url nicht konfiguriert und relative Url gefunden."
        }

        $finalUrl = "$($wikiBaseUrl.TrimEnd('/'))/$($relativePath.TrimStart('/'))";
    }

    if ($jsonContent) {
        return Invoke-WebRequest $finalUrl -UseBasicParsing -Headers @{ Authorization = $wikiBasicAuthValue } -Method $method -Body ([System.Text.Encoding]::UTF8.GetBytes($jsonContent)) -ContentType "application/json" -proxy "http://webproxy.dc.migros.ch:9099"
    }
    if ($attachFile) {
        # Source: https://hochwald.net/upload-file-powershell-invoke-restmethod/
		# The boundary is essential - Trust me, very essential
		$boundary = [Guid]::NewGuid().ToString()
		$bodyStart = @"
--$boundary
Content-Disposition: form-data; name="file"; filename="$(Split-Path -Leaf -Path $attachFile)"
Content-Type: application/octet-stream


"@
		$bodyEnd = @"

--$boundary--
"@
		$requestInFile = "$PSScriptRoot\payload.bin"
        $fileStream = (New-Object -TypeName 'System.IO.FileStream' -ArgumentList ($requestInFile, [IO.FileMode]'Create', [IO.FileAccess]'Write'))

        try
        {
            # The Body start
            $bytes = [Text.Encoding]::UTF8.GetBytes($bodyStart)
            $fileStream.Write($bytes, 0, $bytes.Length)

            # The original File
            $bytes = [IO.File]::ReadAllBytes($attachFile)
            $fileStream.Write($bytes, 0, $bytes.Length)

            # Append the end of the body part
            $bytes = [Text.Encoding]::UTF8.GetBytes($bodyEnd)
            $fileStream.Write($bytes, 0, $bytes.Length)
        }
        finally
        {
            # End the Stream to close the file
            $fileStream.Close()
        }

        # Make it multipart, this is the magic part...
        $contentType = 'multipart/form-data; boundary={0}' -f $boundary

        $response = Invoke-RestMethod $finalUrl -UseBasicParsing -Headers @{ Authorization = $wikiBasicAuthValue; "X-Atlassian-Token" = "no-check" } -Method POST -InFile $requestInFile -ContentType $contentType -proxy "http://webproxy.dc.migros.ch:9099"
    }
    if ($outFile) {
        return Invoke-WebRequest $finalUrl -UseBasicParsing -Headers @{ Authorization = $wikiBasicAuthValue } -OutFile $outFile -proxy "http://webproxy.dc.migros.ch:9099"
    }
    return Invoke-WebRequest $finalUrl -UseBasicParsing -Headers @{ Authorization = $wikiBasicAuthValue } -proxy "http://webproxy.dc.migros.ch:9099"
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
    [System.IO.File]::WriteAllText("$PSScriptRoot\payload.json", $jsonContent)

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
    [System.IO.File]::WriteAllText("$PSScriptRoot\payload.json", $jsonContent)

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

#get page source
$sourcePageData = Invoke-WikiRequest "/rest/api/content/$($pageid)?expand=body.storage,version,space"
$sourcePageDataContent = (ConvertFrom-Json $sourcePageData.Content)
$pageVersion = $sourcePageDataContent.version.number
$pageTitle = $sourcePageDataContent.title
$spaceKey = $sourcePageDataContent.space.key
$sourcePageBody = $sourcePageDataContent.body.storage.value
[System.IO.File]::WriteAllText("$PSScriptRoot\sourcepage.xml", (WrapInHtmlContainer($sourcePageBody)))

$sourcePageBodyXml = [System.Xml.XmlDocument]::new()
$sourcePageBodyXml.LoadXml((Get-Content "$PSScriptRoot\sourcepage.xml" -Raw))

$childPageNumber = 1
$h1Nodes = $sourcePageBodyXml.DocumentElement.SelectNodes("//h1")
$h1Nodes | ForEach-Object {
    # Leeres Dokument erstellen als Ziel
    $newDocument = [System.Xml.XmlDocument]::new()
    $newDocument.LoadXml((WrapInHtmlContainer('')))

    $nodeToTransfer = $_
    $nsmgr = [System.Xml.XmlNamespaceManager]::new($sourcePageBodyXml.NameTable);
    $nsmgr.AddNamespace("ac", "atlassian-confluence");
    $nsmgr.AddNamespace("ri", "atlassian-confluence-ri");
    $nodeToRemove = $nodeToTransfer.SelectSingleNode("ac:structured-macro[@ac:name='anchor']", $nsmgr);
    while ($nodeToRemove) {
        $removed = $nodeToTransfer.RemoveChild($nodeToRemove);
        $nodeToRemove = $nodeToTransfer.SelectSingleNode("ac:structured-macro[@ac:name='anchor']", $nsmgr);
    }
    $newPageTitle = "$($childPageNumber.ToString('00')). $($_.InnerText) ($pageid)"
    Write-Host "Extrahiere Abschnitt '$newPageTitle'..."

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
    
    $newPageResult = Write-NewConfluencePage -spaceKey $spaceKey -parentPageId $pageid -title $newPageTitle -content $newDocument.DocumentElement.InnerXml
    $newPageId = (ConvertFrom-Json $newPageResult.Content).id

    # copy attachments
    $nsmgr = [System.Xml.XmlNamespaceManager]::new($newDocument.NameTable);
    $nsmgr.AddNamespace("ri", "atlassian-confluence-ri");
    $attachmentsToCopy = $newDocument.SelectNodes("//ri:attachment/@ri:filename", $nsmgr)
    if ($attachmentsToCopy.Count -gt 0) {
        $attachmentsToCopy | ForEach-Object {
            $attachmentName = $_.Value
            Write-Host "  - lade Attachment '$attachmentName'"
            $attachmentDownloaded = Invoke-WikiRequest "/download/attachments/$pageid/$attachmentName" -outFile $PSScriptRoot\$attachmentName
            Write-Host "  - schreibe Attachment '$attachmentName'"
            $attachmentUploaded = Invoke-WikiRequest "/rest/api/content/$newPageId/child/attachment" -method POST -attachFile $PSScriptRoot\$attachmentName
            Write-Host "  Attachment '$attachmentName' übertragen"
            Remove-Item $PSScriptRoot\$attachmentName
        }
    }
    $childPageNumber = $childPageNumber+1
}

$remainingContent = $sourcePageBodyXml.DocumentElement.InnerXml
$updatedPage = Write-ExistingConfluencePage -pageId $pageid -spaceKey $spacekey -currentVersion $pageVersion -title $pageTitle -content $remainingContent
