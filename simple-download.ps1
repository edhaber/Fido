function Get-RandomDate()
{
	[DateTime]$Min = "1/1/2008"
	[DateTime]$Max = [DateTime]::Now

	$RandomGen = new-object random
	$RandomTicks = [Convert]::ToInt64( ($Max.ticks * 1.0 - $Min.Ticks * 1.0 ) * $RandomGen.NextDouble() + $Min.Ticks * 1.0 )
	$Date = new-object DateTime($RandomTicks)
	return $Date.ToString("yyyyMMdd")
}

function Error([string]$ErrorMessage)
{
	Write-Host Error: $ErrorMessage
	$script:ExitCode = $script:Stage--
}

$QueryLocale = "en-US"
$SessionId = [guid]::NewGuid()
# @(
# 			"20H1 (Build 19041.264 - 2020.05)",
# 			@("Windows 10 Home/Pro", 1626),
# 			@("Windows 10 Education", 1625),
# 			@("Windows 10 Home China ", ($zh + 1627))
# 		),
# 		@(
# 			"19H2 (Build 18363.418 - 2019.11)",
# 			@("Windows 10 Home/Pro", 1429),
# 			@("Windows 10 Education", 1431),
# 			@("Windows 10 Home China ", ($zh + 1430))
# 		),
# 		@(
# 			"19H1 (Build 18362.356 - 2019.09)",
# 			@("Windows 10 Home/Pro", 1384),
# 			@("Windows 10 Education", 1386),
# 			@("Windows 10 Home China ", ($zh + 1385))
# 		),
$WindowsVersion = "Windows10ISO"
$skuID = 8143
$ProductEdition = 1626
$Language = "English"
$Is64 = $true
$ExitCode = 100
$Locale = "en-US"
$DFRCKey = "HKLM:\Software\Policies\Microsoft\Internet Explorer\Main\"
$DFRCName = "DisableFirstRunCustomize"
$DFRCAdded = $False
$RequestData = @{}
$RequestData["GetLangs"] = @("a8f8f489-4c7f-463a-9ca6-5cff94d8d041", "getskuinformationbyproductedition" )
$RequestData["GetLinks"] = @("cfa9e580-a81e-4a4b-a846-7b21bf4e2e5b", "GetProductDownloadLinksBySku" )
# Create a semi-random Linux User-Agent string
$FirefoxVersion = Get-Random -Minimum 30 -Maximum 60
$FirefoxDate = Get-RandomDate
$UserAgent = "Mozilla/5.0 (X11; Linux i586; rv:$FirefoxVersion.0) Gecko/$FirefoxDate Firefox/$FirefoxVersion.0"
#endregion

$url = "https://www.microsoft.com/" + $QueryLocale + "/api/controls/contentinclude/html"
$url += "?pageId=" + $RequestData["GetLangs"][0]
$url += "&host=www.microsoft.com"
$url += "&segments=software-download," + $WindowsVersion
$url += "&query=&action=" + $RequestData["GetLangs"][1]
$url += "&sessionId=" + $SessionId
$url += "&productEditionId=" + [Math]::Abs($ProductEdition)
$url += "&sdVersion=2"
Write-Host Querying $url

$r = Invoke-WebRequest -UseBasicParsing -UserAgent $UserAgent -WebSession $Session $url 

#https://www.microsoft.com/en-us/api/controls/contentinclude/html?pageId=cfa9e580-a81e-4a4b-a846-7b21bf4e2e5b&
#host=www.microsoft.com&segments=software-download%2cwindows10ISO&query=&action=
#GetProductDownloadLinksBySku&sessionId=31c37d42-dd8d-47b7-b8ab-762016cdb669&skuId=8143&language=English&sdVersion=2
$url = "https://www.microsoft.com/" + $QueryLocale + "/api/controls/contentinclude/html"
$url += "?pageId=" + $RequestData["GetLinks"][0]
$url += "&host=www.microsoft.com"
$url += "&segments=software-download," + $WindowsVersion
$url += "&query=&action=" + $RequestData["GetLinks"][1]
$url += "&sessionId=" + $SessionId
$url += "&skuId=" + $skuID
$url += "&language=" + $Language
$url += "&sdVersion=2"
Write-Host Querying $url

$i = 0
$SelectedIndex = 0
$array = @()
try {
    $Is64 = [Environment]::Is64BitOperatingSystem
    $r = Invoke-WebRequest -UseBasicParsing -UserAgent $UserAgent -WebSession $Session $url
    
    $HTMLr = New-Object -Com "HTMLFile"
    $HTMLr.IHTMLDocument2_write($r.Content)
    $html = $($HTMLr.all.tags("input")).outerHTML
    # Need to fix the HTML and JSON data so that it is well-formed
    $html = $html.Replace("class=product-download-hidden", "")
    $html = $html.Replace("type=hidden", "")
    $html = $html.Replace(">", "/>")
    $html = $html.Replace("IsoX86", """x86""")
    $html = $html.Replace("IsoX64", """x64""")
    $html = "<inputs>" + $html + "</inputs>"
    $xml = [xml]$html
    foreach ($var in $xml.inputs.input) {
        $json = $var.value | ConvertFrom-Json;
        if ($json) {
            if (($Is64 -and $json.DownloadType -eq "x64") -or (-not $Is64 -and $json.DownloadType -eq "x86")) {
                $SelectedIndex = $i
            }
            $array += @(New-Object PsObject -Property @{ Type = $json.DownloadType; Link = $json.Uri })
            $i++
        }
    }
    if ($array.Length -eq 0) {
        Throw-Error -Req $r -Alt "Could not retreive ISO download links"
    }
} catch {
    Write-Host $_.Exception
    Error($_.Exception.Message)
    return
}
$array[$SelectedIndex].Link