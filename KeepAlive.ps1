# Ping Dovetail tenants to keep them alive

# TODO:
# Figure out a way to keep the general-api app alive
# Chad: I think Site24x7 hitting the login page isn't enough now.  I noticed that there was a hit to the AutoMapperConfiguration that seemed to take awhile. So I think there's a big hit the first time you open a query grid. Not just the setup page, but you  need to actually hit a grid. 


# Tenants & Portals
$tenants = @('https://default.lclhst.io','https://dev.dovetailnow.com')
$portals = @('http://default.lclhst.io/esp','https://dev.portal.demo.dovetailnow.com')

# How long to sleep betweek keep-alives (15 is a good number, as IIS usually has a 20 minute app timeout)
$sleepTimeInMinutes = 15

# How many keep alive loops to do?
# 4 loops of 15 minutes each, would keep the app alive for the next hour
# 8 loops of 15 minutes each, would keep the app alive for the next 2 hours
$numberOfLoops = 4


### Functions

function Start-Sleep($seconds) {
    $doneDT = (Get-Date).AddSeconds($seconds)
    while($doneDT -gt (Get-Date)) {
        $secondsLeft = $doneDT.Subtract((Get-Date)).TotalSeconds
        $percent = ($seconds - $secondsLeft) / $seconds * 100
        Write-Progress -Activity "Sleeping" -Status "Sleeping..." -SecondsRemaining $secondsLeft -PercentComplete $percent
        [System.Threading.Thread]::Sleep(500)
    }
    Write-Progress -Activity "Sleeping" -Status "Sleeping..." -SecondsRemaining 0 -Completed
}


function Ping-Tenants(){	
	# Setup
	$WebClient = New-Object Net.WebClient 

	foreach ($tenant in $tenants) {
		
		write-host "Pinging: $tenant"

		# Agent Legacy
		$Source = $WebClient.DownloadString($tenant + "/agent/health.aspx")

		# Agent SPA
		$Source = $WebClient.DownloadString($tenant + "/agent/public/case-survey/123/456")

		# API Docs
		$Source = $WebClient.DownloadString($tenant + "/api/doc")

		# Legacy API
		$Source = $WebClient.DownloadString($tenant + "/api/health.aspx")

		# Portal Configs
		# Unsure if the health-check endpoint really spins up the app completely.  
		$Source = $WebClient.DownloadString($tenant + "/api/v3/portal-configs/health-check");

		# This will 401 (unauthorized). I have the doubts whether this really spins up the app. Blurg. Regardless, just fail silently
		try {
			$Source = $WebClient.DownloadString($tenant + "/agent/api/portal-configs");
			}
		catch {}

		# General API

		# This will 401 (unauthorized). I have the doubts whether this really spins up the app. Blurg. Regardless, just fail silently
		try {
			$Source = $WebClient.DownloadString($tenant + "/api/assets/00000000-0000-0000-0000-000000000000");	
			}	
		catch {}

		# This works locally, but not in AWS (because general-api isn't exposed publicly, and we don't have a redirect rule, as it doesn't need to be publicly exposed)
		if ($tenant -eq 'https://default.lclhst.io'){
			$Source = $WebClient.DownloadString($tenant + "/general-api/health-check")		
		}

	}

	foreach ($portal in $portals) {
		write-host "Pinging Portal: $portal"
		$Source = $WebClient.DownloadString($portal + "/login")
	}

}


### Main

$loopCounter = 1 
$sleepTimeInSeconds = 60 * $sleepTimeInMinutes 

DO {
	"Starting Loop $loopCounter of $numberOfLoops"
	Ping-Tenants
	Start-Sleep $sleepTimeInSeconds
	$loopCounter++
} While ($loopCounter -le $numberOfLoops)

