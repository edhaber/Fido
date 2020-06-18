function Get-RandomDate()
{
	[DateTime]$Min = "1/1/2008"
	[DateTime]$Max = [DateTime]::Now

	$RandomGen = new-object random
	$RandomTicks = [Convert]::ToInt64( ($Max.ticks * 1.0 - $Min.Ticks * 1.0 ) * $RandomGen.NextDouble() + $Min.Ticks * 1.0 )
	$Date = new-object DateTime($RandomTicks)
	return $Date.ToString("yyyyMMdd")
}

function Throw-Error([object]$Req, [string]$Alt)
{
	$Err = $(GetElementById -Request $r -Id "errorModalMessage").innerText
	if (-not $Err) {
		$Err = $Alt
	} else {
		$Err = [System.Text.Encoding]::UTF8.GetString([byte[]][char[]]$Err)
	}
	throw $Err
}

# Some PowerShells don't have Microsoft.mshtml assembly (comes with MS Office?)
# so we can't use ParsedHtml or IHTMLDocument[2|3] features there...
function GetElementById([object]$Request, [string]$Id)
{
	try {
		return $Request.ParsedHtml.IHTMLDocument3_GetElementByID($Id)
	} catch {
		return $Request.AllElements | ? {$_.id -eq $Id}
	}
}

function Select-Language([string]$LangName)
{
	# Use the system locale to try select the most appropriate language
	[string]$SysLocale = [System.Globalization.CultureInfo]::CurrentUICulture.Name
	if (($SysLocale.StartsWith("ar") -and $LangName -like "*Arabic*") -or `
		($SysLocale -eq "pt-BR" -and $LangName -like "*Brazil*") -or `
		($SysLocale.StartsWith("ar") -and $LangName -like "*Bulgar*") -or `
		($SysLocale -eq "zh-CN" -and $LangName -like "*Chinese*" -and $LangName -like "*simp*") -or `
		($SysLocale -eq "zh-TW" -and $LangName -like "*Chinese*" -and $LangName -like "*trad*") -or `
		($SysLocale.StartsWith("hr") -and $LangName -like "*Croat*") -or `
		($SysLocale.StartsWith("cz") -and $LangName -like "*Czech*") -or `
		($SysLocale.StartsWith("da") -and $LangName -like "*Danish*") -or `
		($SysLocale.StartsWith("nl") -and $LangName -like "*Dutch*") -or `
		($SysLocale -eq "en-US" -and $LangName -eq "English") -or `
		($SysLocale.StartsWith("en") -and $LangName -like "*English*" -and ($LangName -like "*inter*" -or $LangName -like "*ingdom*")) -or `
		($SysLocale.StartsWith("et") -and $LangName -like "*Eston*") -or `
		($SysLocale.StartsWith("fi") -and $LangName -like "*Finn*") -or `
		($SysLocale -eq "fr-CA" -and $LangName -like "*French*" -and $LangName -like "*Canad*") -or `
		($SysLocale.StartsWith("fr") -and $LangName -eq "French") -or `
		($SysLocale.StartsWith("de") -and $LangName -like "*German*") -or `
		($SysLocale.StartsWith("el") -and $LangName -like "*Greek*") -or `
		($SysLocale.StartsWith("he") -and $LangName -like "*Hebrew*") -or `
		($SysLocale.StartsWith("hu") -and $LangName -like "*Hungar*") -or `
		($SysLocale.StartsWith("id") -and $LangName -like "*Indones*") -or `
		($SysLocale.StartsWith("it") -and $LangName -like "*Italia*") -or `
		($SysLocale.StartsWith("ja") -and $LangName -like "*Japan*") -or `
		($SysLocale.StartsWith("ko") -and $LangName -like "*Korea*") -or `
		($SysLocale.StartsWith("lv") -and $LangName -like "*Latvia*") -or `
		($SysLocale.StartsWith("lt") -and $LangName -like "*Lithuania*") -or `
		($SysLocale.StartsWith("ms") -and $LangName -like "*Malay*") -or `
		($SysLocale.StartsWith("nb") -and $LangName -like "*Norw*") -or `
		($SysLocale.StartsWith("fa") -and $LangName -like "*Persia*") -or `
		($SysLocale.StartsWith("pl") -and $LangName -like "*Polish*") -or `
		($SysLocale -eq "pt-PT" -and $LangName -eq "Portuguese") -or `
		($SysLocale.StartsWith("ro") -and $LangName -like "*Romania*") -or `
		($SysLocale.StartsWith("ru") -and $LangName -like "*Russia*") -or `
		($SysLocale.StartsWith("sr") -and $LangName -like "*Serbia*") -or `
		($SysLocale.StartsWith("sk") -and $LangName -like "*Slovak*") -or `
		($SysLocale.StartsWith("sl") -and $LangName -like "*Slovenia*") -or `
		($SysLocale -eq "es-ES" -and $LangName -eq "Spanish") -or `
		($SysLocale.StartsWith("es") -and $Locale -ne "es-ES" -and $LangName -like "*Spanish*") -or `
		($SysLocale.StartsWith("sv") -and $LangName -like "*Swed*") -or `
		($SysLocale.StartsWith("th") -and $LangName -like "*Thai*") -or `
		($SysLocale.StartsWith("tr") -and $LangName -like "*Turk*") -or `
		($SysLocale.StartsWith("uk") -and $LangName -like "*Ukrain*") -or `
		($SysLocale.StartsWith("vi") -and $LangName -like "*Vietnam*")) {
		return $True
	}
	return $False
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
$ProductEdition = 1429
$LanguageName = "English"

$Language = @{}
$Is64 = $true
$ExitCode = 100
$Locale = "en-US"
# $DisableFirstRunCustomize = $true
# $DFRCKey = "HKLM:\Software\Policies\Microsoft\Internet Explorer\Main\"
# $DFRCName = "DisableFirstRunCustomize"
# $DFRCAdded = $False
$RequestData = @{}
$RequestData["GetLangs"] = @("a8f8f489-4c7f-463a-9ca6-5cff94d8d041", "getskuinformationbyproductedition" )
$RequestData["GetLinks"] = @("cfa9e580-a81e-4a4b-a846-7b21bf4e2e5b", "GetProductDownloadLinksBySku" )
# Create a semi-random Linux User-Agent string
$FirefoxVersion = Get-Random -Minimum 30 -Maximum 60
$FirefoxDate = Get-RandomDate
$UserAgent = "Mozilla/5.0 (X11; Linux i586; rv:$FirefoxVersion.0) Gecko/$FirefoxDate Firefox/$FirefoxVersion.0"
# Localization
$EnglishMessages = "en-US|Version|Release|Edition|Language|Architecture|Download|Continue|Back|Close|Cancel|Error|Please wait...|" +
	"Download using a browser|Temporarily banned by Microsoft for requesting too many downloads - Please try again later...|" +
	"PowerShell 3.0 or later is required to run this script.|Do you want to go online and download it?"
[string[]]$English = $EnglishMessages.Split('|')
[string[]]$Localized = $null
if ($LocData -and (-not $LocData.StartsWith("en-US"))) {
	$Localized = $LocData.Split('|')
	if ($Localized.Length -ne $English.Length) {
		Write-Host "Error: Missing or extra translated messages provided ($($Localized.Length)/$($English.Length))"
		exit 101
	}
	$Locale = $Localized[0]
}
$QueryLocale = $Locale
#endregion


# If asked, disable IE's first run customize prompt as it interferes with Invoke-WebRequest
# if ($DisableFirstRunCustomize) {
# 	try {
# 		# Only create the key if it doesn't already exist
# 		Get-ItemProperty -Path $DFRCKey -Name $DFRCName
# 	} catch {
# 		if (-not (Test-Path $DFRCKey)) {
# 			New-Item -Path $DFRCKey -Force | Out-Null
# 		}
# 		Set-ItemProperty -Path $DFRCKey -Name $DFRCName -Value 1
# 		$DFRCAdded = $True
# 	}
# }

$url = "https://www.microsoft.com/" + $QueryLocale + "/api/controls/contentinclude/html"
$url += "?pageId=" + $RequestData["GetLangs"][0]
$url += "&host=www.microsoft.com"
$url += "&segments=software-download," + $WindowsVersion
$url += "&query=&action=" + $RequestData["GetLangs"][1]
$url += "&sessionId=" + $SessionId
$url += "&productEditionId=" + [Math]::Abs($ProductEdition)
$url += "&sdVersion=2"
#Write-Host Querying $url

#$r = Invoke-WebRequest -UseBasicParsing -UserAgent $UserAgent -WebSession $Session $url 


$array = @()
$i = 0
$SelectedIndex = 0
try {
    $r = Invoke-WebRequest -UserAgent $UserAgent -WebSession $Session $url
    # Go through an XML conversion to keep all PowerShells happy...
    if (-not $($r.AllElements | ? {$_.id -eq "product-languages"})) {
        throw "Unexpected server response"
    }
    $html = $($r.AllElements | ? {$_.id -eq "product-languages"}).InnerHTML
    $html = $html.Replace("selected value", "value")
    $html = $html.Replace("&", "&amp;")
    $html = "<options>" + $html + "</options>"
    $xml = [xml]$html
    foreach ($var in $xml.options.option) {
        $json = $var.value | ConvertFrom-Json;
        if ($json) {
            $array += @(New-Object PsObject -Property @{ DisplayLanguage = $var.InnerText; Language = $json.language; Id = $json.id })
            #Write-Output $json.language
            if ($json.language -eq $LanguageName) {
                $SelectedIndex = $i
            }
            $i++
        }
    }
    if ($array.Length -eq 0) {
        Throw-Error -Req $r -Alt "Could not parse languages"
    }
} catch {
    Error($_.Exception.Message)
    break
}
$Language = $array[$SelectedIndex]

#https://www.microsoft.com/en-us/api/controls/contentinclude/html?pageId=cfa9e580-a81e-4a4b-a846-7b21bf4e2e5b&
#host=www.microsoft.com&segments=software-download%2cwindows10ISO&query=&action=
#GetProductDownloadLinksBySku&sessionId=31c37d42-dd8d-47b7-b8ab-762016cdb669&skuId=8143&language=English&sdVersion=2
$url = "https://www.microsoft.com/" + $QueryLocale + "/api/controls/contentinclude/html"
$url += "?pageId=" + $RequestData["GetLinks"][0]
$url += "&host=www.microsoft.com"
$url += "&segments=software-download," + $WindowsVersion
$url += "&query=&action=" + $RequestData["GetLinks"][1]
$url += "&sessionId=" + $SessionId
$url += "&skuId=" + $Language.Id
$url += "&language=" + $Language.Language
$url += "&sdVersion=2"

#Write-Host Querying $url

$i = 0
$SelectedIndex = 0
$array = @()
try {
    $Is64 = [Environment]::Is64BitOperatingSystem
    $r = Invoke-WebRequest -UserAgent $UserAgent -WebSession $Session $url
    if (-not $($r.AllElements | ? {$_.id -eq "expiration-time"})) {
        Throw-Error -Req $r -Alt $English[14]
    }
    $html = $($r.AllElements | ? {$_.tagname -eq "input"}).outerHTML

    # $HTMLr = New-Object -Com "HTMLFile"
    # $HTMLr.IHTMLDocument2_write($r.Content)
    # $html = $($HTMLr.all.tags("input")).outerHTML
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
    Error($_.Exception.Message)
    return
}
$array[$SelectedIndex].Link