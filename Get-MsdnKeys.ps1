# It might work in 4, but I'm not testing there. Lower you'll have to tweak code
#requires -Version 5.0
param(
    # Your Live ID for MSDN login
    [Parameter(Mandatory)]
    [PSCredential]
    [System.Management.Automation.CredentialAttribute()]
    $Credential,

    # Pick a browser to use. Defaults to Firefox (which doesn't seem to require an external Driver file) 
    # Works for sure with Firefox, Chrome, and Edge, and PhantomJS
    # The edge driver must be manually installed from http://go.microsoft.com/fwlink/?LinkId=619687
    # The PhantomJS driver must be manually installed from http://phantomjs.org/download.html
    # The IE driver on NuGet did not work for me.
    # And I can't be bothered to mess with Opera or any of the others
    [ValidateSet("Chrome", "Firefox", "Edge", "PhantomJS")] # 
    $Browser = "Firefox"
)


function Install-Selenium {
    param(
        [Parameter(Mandatory)]
        [ValidateSet("Chrome", "Firefox", "InternetExplorer", "Edge", "Opera", "PhantomJS")] # 
        $Browser = "Firefox"
    )

    $Package = switch($Browser) {
        "Edge" { "MicrosoftWebDriver" }
        "InternetExplorer" { "IEDriver" }
        "PhantomJS" { "PhantomJS" }
        default { "${_}Driver" }
    }

    if(!("OpenQA.Selenium.By" -as [type])) {
        if(!($WebDriverPath = Resolve-Path "$PSScriptRoot\Selenium.WebDriver.*\lib\net40" -ErrorActionPreference SilentlyContinue)) {
            # Install Selenium from NuGet
            Install-Package Selenium.WebDriver -Destination $PSScriptRoot
            $WebDriverPath = Resolve-Path "$PSScriptRoot\Selenium.WebDriver.*\lib\net40"
        }
        # Load Selenium
        Add-Type -Path (Join-Path $WebDriverPath WebDriver.dll)
    }

    # If we need a driver ...
    if(($Browser -ne "Firefox") -and !(Get-Command "${Package}*.exe")) {
        # Chrome seems to be the only one we can auto-fetch from NuGet
        # The others may have packages on NuGet, but those did not work for me
        if($Browser -eq "Chrome") {
            Install-Package Selenium.WebDriver.${Package} -Destination $PSScriptRoot
            $Env:Path += ";" + ((Resolve-Path "scripts\Selenium.WebDriver.*Driver.*\driver\") -join ";")
        }
    }

    $ErrorActionPreference = "Stop"
}


function Get-Selenium {
    #.Synopsis
    #   Get a Selenium automation driver for the specified browser
    param(
        [Parameter(Mandatory)]
        [ValidateSet("Chrome", "Firefox", "InternetExplorer", "Edge", "Opera", "PhantomJS")] # 
        $Browser = "Firefox"
    )

    $Driver = switch($Browser) {
        "InternetExplorer" { "IE" }
        default { $_ }
    }

    if(!("OpenQA.Selenium.${Driver}.${Browser}Driver" -as [type])) {
        Install-Selenium $Browser
    }

    return ($global:Selenium = New-Object OpenQA.Selenium.${Driver}.${Browser}Driver)
}


$Selenium = Get-Selenium $Browser

# TODO: See if it's enough to just go to login.live.com
$msdn = "https://login.live.com/login.srf?wa=wsignin1.0&wreply=https%3a%2f%2fmsdn.microsoft.com%2fen-us%2fsubscriptions%2fdownloads%2f"

$Selenium.Navigate().GoToUrl($msdn)

Start-Sleep 1
# Sometimes it remembers my name
if(($name = $Selenium.FindElementByCssSelector('input[type="email"]')) -and $name.Displayed) {
    Write-Verbose "Sending UserName"
    $name.SendKeys($Credential.UserName)
}

# But if the password box isn't there, we were probably already logged in
$pass = $Selenium.FindElementByCssSelector('input[type="password"]')
$pass.Clear() # sometimes it's pre-populated if your browser stores it...
$pass.SendKeys($Credential.GetNetworkCredential().Password) 

$send = $Selenium.FindElementByCssSelector('input[type="submit"]')
$send.Submit()

$count = 0
while($Selenium.Url -notmatch "^https://msdn.microsoft.com/.*/downloads/") {
    Start-Sleep -milli 500
    if(2 -lt $count++) {
        if($Selenium.Url.StartsWith("https://login.live.com/ppsecure/post.srf")) {
            if($otc = $Selenium.FindElementByCssSelector('input[name="otc"]')) {
                $code = Read-Host "We need your 2FA one-time code"
                $otc.SendKeys($code)
                $otc.Submit()
            }
            break
        } else {
            Write-Warning "We don't seem to have arrived at the downloads page. Current Url: $($Selenium.Url)"
            break
        } 
    }
}
if($Selenium.Url.StartsWith("https://login.live.com/ppsecure/post.srf")) {
    if($otc = $Selenium.FindElementByCssSelector('input[name="otc"]')) {
        $code = Read-Host "We need your 2FA one-time code"
        $otc.SendKeys($code)
        $otc.Submit()
    }
}

if($Selenium.Manage().Cookies.AllCookies.Count -eq 0) {
    throw "Couldn't get authentication cookie from $Browser"
}

# Precreate a session object
$response  = Invoke-WebRequest login.live.com -SessionVariable Jar
# Fill it with nice warm cookies (ignoring expiration dates)
$Selenium.Manage().Cookies.AllCookies | Select-Object Name, Value, Domain, Secure | % { $Jar.Cookies.Add([Net.Cookie]$_) }
# Trade it for the MSDN keys
[xml](Microsoft.PowerShell.Utility\Invoke-WebRequest https://msdn.microsoft.com/en-us/subscriptions/securejson/getallexportkeys?brand=msdn -WebSession $Jar).Content