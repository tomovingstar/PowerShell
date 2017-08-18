<#
.Synopsis
	This Script shows menu in PowerShell window with 4 available options, namely:
	"Initial Login" (i) is to open URLs and start executables at the beginning of the workday.
		URLs and executables are at the beginning of the script and can be edited with Notepad.
	"Display Form" (d) is to show a pop-up window with a list of documents that can be copied to Windows clipboard.
		Documents are stored in DocumentBase.xml file (loaded by script from the same directory) and can be edited with Notepad.
	"Start Alarm" (s) starts voice alarm at selected time settings.
		Time settings are at the beginning of the script and can be edited with Notepad.
	"Quit" (q) quits script.
	
#>

#Enter alarm set values for your shift (min)
$ShiftStart = 8*60 + 55
$LunchStart = 13*60 + 0
$LunchEnd = 14*60 + 0
$ShiftEnd = 17*60 + 55
#Enter URLs to open
[array]$URLs = @(
"https://www.google.com/",
"https://www.microsoft.com/en-us/",
"https://www.yahoo.com/"
)

#Enter executables to start
[array]$Apps = @(
"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook 2016.lnk"
)
[array]$Messages = @(
"Gentlemen, could you please update last contact records on your cases? Thank you.",
"Gentlemen, could you please update wait state on your cases. Thank you.",
"Gentlemen, could you please schedule phone calls on your cases. Thank you."
)
#---------------
$CurrentPath = split-path $SCRIPT:MyInvocation.MyCommand.Path -parent
[xml]$myData = Get-Content $CurrentPath/DocumentBase.xml
$SetTimes = @($ShiftStart, $LunchStart, $LunchEnd, $ShiftEnd)
	
function fOutput {

	$title = 'Productivity Center'
	$message = 'Press key listed below and hit "Enter" for the following action:'

	$initial = New-Object System.Management.Automation.Host.ChoiceDescription "&Initial Login", "Initial Login"
	$showform = New-Object System.Management.Automation.Host.ChoiceDescription "&Display DocBase", "DocBase"
	$startalarm = New-Object System.Management.Automation.Host.ChoiceDescription "&Start Alarm", "Start Alarm"
	$quit = New-Object System.Management.Automation.Host.ChoiceDescription "&Quit", "Quit"

	$options = [System.Management.Automation.Host.ChoiceDescription[]]($initial, $showform, $startalarm, $quit)
	$result = $host.ui.PromptForChoice($title, $message, $options, 2) 

	switch ($result) {
			0 { fInitial }
			1 { fShowform }
			2 { fStartalarm }
			3 { "quitting..." }
	}
}

function fInitial {
	
		Show-ControlPanelItem -Name "Windows Update"
		Show-ControlPanelItem -Name "Sound"
		
		$ie = New-Object -ComObject InternetExplorer.Application
		$ie.Navigate2($URLs[0]);
		[int]$StopAt = $URLs.Length - 3
				
		For($i = 1; $i -le $StopAt; $i++){
			$ie.Navigate2($URLs[$i], 0x1000)
		}
		
		$ie.Visible = $true		
		$ie = $Null

		ForEach ($App In $Apps) {
			& $App
		}
		fOutput
}

function fShowForm {
	Start-Job -Name ShowForm -ScriptBlock {

		Add-Type -Assembly System.Windows.Forms
		Add-Type -AssemblyName System.Drawing 
			
		$DocBase = $using:myData

		$label = New-Object System.Windows.Forms.Label 
		$label.Location = New-Object System.Drawing.Size(10,5)  
		$label.Size = New-Object System.Drawing.Size(280,20) 
		$label.AutoSize = $true
		$label.Text = "Select item to copy doc to clipboard"
		
		$listBox = New-Object System.Windows.Forms.ListBox 
		$listBox.Location = New-Object System.Drawing.Size(5,25) 
		$listBox.Size = New-Object System.Drawing.Size(190,230) 
		
		foreach ($strTitle in $DocBase.DocumentBase.Document.Title){
    			[void] $listBox.Items.Add($strTitle)
    	}
		$listBox.Add_Click({ $DocBase.DocumentBase.Document[$listBox.SelectedIndex].Content | clip.exe })
		
		$cancelButton = New-Object System.Windows.Forms.Button 
		$cancelButton.Location = New-Object System.Drawing.Size(115,255) 
		$cancelButton.Size = New-Object System.Drawing.Size(75,25) 
		$cancelButton.Text = "Cancel"
		$cancelButton.Add_Click({ $form.Close() }) 
		
		$form = New-Object System.Windows.Forms.Form  
		$form.Text = "DocBase"
		$form.Size = New-Object System.Drawing.Size(215,320) 
		$form.FormBorderStyle = 'FixedSingle'
		$form.StartPosition = "CenterScreen"
		$form.AutoSizeMode = 'GrowAndShrink'
		$form.Topmost = $true
		$form.CancelButton = $cancelButton
		$form.ShowInTaskbar = $true
		  
		$form.Controls.Add($label) 
		$form.Controls.Add($textBox) 
		$form.Controls.Add($cancelButton) 
		$form.Controls.Add($listBox) 
		$form.Topmost = $true
		
		$form.Add_Shown({$form.Activate()}) 
		$form.ShowDialog()
	}
	fOutput
}

function fStartAlarm {
	Start-Job -Name StartAlarm -ScriptBlock {
		
		$myMessages = $using:Messages
		$CurTime=Get-Date
		
		foreach ($SetTime in $using:SetTimes) {
			if ( ($CurTime.Hour * 60 + $CurTime.Minute) -ge $SetTime) { continue }

			do {

				$CurTime=Get-Date
				Start-Sleep -Seconds 5

			} Until ( ($CurTime.Hour * 60 + $CurTime.Minute) -ge $SetTime)

			$ie = New-Object -ComObject InternetExplorer.Application
			$ie.Navigate($using:URLs[-2])
			$ie.Visible = $true
			$ie = $Null

			switch ($SetTime) { 
					$using:SetTimes[-1] {
					$ie = New-Object -ComObject InternetExplorer.Application
					$ie.Navigate($using:URLs[-1])
					$ie.Visible = $true } 
					default {"no default"}
			}
			
			Add-Type -AssemblyName System.speech
			$Speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
			$InstalledVoices = $Speak.GetInstalledVoices() | ForEach-Object { $_.VoiceInfo }
			$InstalledVoices | ForEach-Object { If($_.Name -eq "Microsoft Hazel Desktop") { $Speak.SelectVoice("Microsoft Hazel Desktop") } }
			$Speak.Speak($myMessages[$Random])
		}
	}
	fOutput
}
fOutput
