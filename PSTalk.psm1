New-Alias SPK Speak
$User = $env:USERPROFILE
$Desktop = $User + "\desktop" 

<#.Synopsis
	The Speak function allows you to use PowerShell to speak. 
.DESCRIPTION
	-Text: Allows you to type the text to speak
	-Volume: Use it to select a specific volume 'from 0 to 100) 
	-Speed: Use it to select a specific speed 'from -10 to 10) 
	-Voice: Use it to choose which language to use
	-Generate: Use it to generate a script on your desktop that will use your configuration
	-Resume: Use to add a short resume about selected options
	
	Speak cmdlet using with all parameters
		speak -Text "Let us play with PowerShell" -volume 60 -speed 3 -voice -generate -resume
	
	You can also use a pipeline with the speak cmdlet as below:
		"Let us play with PowerShell" | speak
		OR
		"Let us play with PowerShell" | spk

.EXAMPLE
	PS Root\> speak -Text "Let us play with PowerShell"
	The command above will tell the following text: "Let us play with PowerShell"
	It will use the default volume and speed values (40 and 0)
	It will use the default selected language

.EXAMPLE
	PS Root\> speak -Text "Let us play with PowerShell" -volume 60 -speed 3
	The command above will tell the following text: "Let us play with PowerShell"
	It will use following volume and speed values (60 and 3)
	It will use the default selected language
	
.EXAMPLE
	PS Root\> speak -Text "Let us play with PowerShell" -volume 60 -speed 3 -voice 
	The command above will tell the following text: "Let us play with PowerShell"
	It will use following volume and speed values (60 and 3)
	It will display available voices on your computer and let you choose which language to use
	
.EXAMPLE
	PS Root\> speak -Text "Let us play with PowerShell" -volume 60 -speed 3 -voice -generate
	The command above will tell the following text: "Let us play with PowerShell"
	It will use following volume and speed values (60 and 3)
	It will display available voices on your computer and let you choose which language to use
	It will generate a script on your desktop, My_Speech_Script.ps1, using your choices

.NOTES
    Author: Damien VAN ROBAEYS - @syst_and_deploy - http://www.systanddeploy.com
#>	
	
Function Speak
{
[CmdletBinding()]
Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true, position=1)]		
        [string] $Text,				
		[Parameter(Mandatory=$False,Position=1)]
        [int] $Speed,
		[Parameter(Mandatory=$False,Position=1)]
        [int] $Volume,	
		[Parameter(Mandatory=$False,Position=1)]
        [switch] $Voice,		
		[Parameter(Mandatory=$False,Position=1)]		
        [switch] $Generate,	
		[Parameter(Mandatory=$False,Position=1)]		
        [switch] $Resume				
      )
    
    Begin
    {			
		Try
			{						
				Add-Type -AssemblyName System.speech
				$Global:Talk = New-Object System.Speech.Synthesis.SpeechSynthesizer
				
				If ($Voice)
					{
						$List_Languages = $Talk.GetInstalledVoices().VoiceInfo	
						$lang_Count = $List_Languages.count
						
						write-host "" 									
						write-host "***********************************************************************" 	
						write-host "			VOICE SELECTION" -foregroundcolor "Green"		
						write-host "***********************************************************************" 
						write-host "" 									
																		
						$Lang_Choice = @{}
						for ($i=1;$i -le $List_Languages.count; $i++) {
							Write-Host "$i. $($List_Languages[$i-1].name)" - "$($List_Languages[$i-1].Culture)" - "$($List_Languages[$i-1].Gender)" -foregroundcolor "Cyan"											
							$Lang_Choice.Add($i,($List_Languages[$i-1].name))
							}

						[int]$MyLang = Read-Host "Please select which language to select to speak your text"	
						write-host "" 										
											
						$Lang = $Lang_Choice.Item($MyLang)					
					}	
			}

		Catch 
			{
				write-host ""		
				write-host ""					
				write-host "***********************************************"	
				write-host "Can not load the System Speech assembly" -foregroundcolor "Yellow"	
				write-host "***********************************************"	
				write-host ""				
				exit
			}		
    }
	

    Process
    {
		
		Add-Type -AssemblyName System.speech
		$Global:Talk = New-Object System.Speech.Synthesis.SpeechSynthesizer			
	
		If (-not $Speed)
			{
				$Talk.Rate = "0"
			}	
		Else
			{
				$Speed = "0"			
				$Talk.Rate = $Speed
			}			
			
		If (-not $Volume)
			{
				$Volume = "40"
				$Talk.Volume = "40"
			}	
		Else
			{
				$Talk.Volume = $Volume
			}			
			
		If (-not $Voice)
			{
				$Lang = ($Talk.voice).name
			}	
	
		$Talk.SelectVoice($Lang)																	
		$Talk.Speak($Text)	
	}

    end
    {
		If ($Generate)
			{
				$Generate_ps1 = "True"				
				$Script_File = "$Desktop\My_Speech_Script.ps1"
				New-Item $Script_File -type file				
				Add-Content $Script_File "#Load assembly"	
				Add-Content $Script_File 'Add-Type -AssemblyName System.speech'
				Add-Content $Script_File '$Talk = New-Object System.Speech.Synthesis.SpeechSynthesizer'

				Add-Content $Script_File ""	
				Add-Content $Script_File "# Set the selectd voice"			
				Add-Content $Script_File ('$Talk.SelectVoice' + "('" + "$Lang" + "')")

				Add-Content $Script_File ""	
				Add-Content $Script_File "# Set the speed value"				
				Add-Content $Script_File ('$Talk.Rate = ' + '"' + "$Speed" + '"')

				Add-Content $Script_File ""	
				Add-Content $Script_File "# Set the volume value"			
				Add-Content $Script_File ('$Talk.Volume = ' + '"' + "$Volume" + '"')
					
				Add-Content $Script_File ""	
				Add-Content $Script_File "# Set the text to speak"			
				Add-Content $Script_File ('$Talk.Speak(' + '"' + "$Text" + '")')				
			}		
		Else
			{
				$Generate_ps1 = "False"
			}		
			
		If ($Resume)
			{						
				write-host "" 				
				write-host "***********************************************************************" 	
				write-host "			VOICE RESUME SELECTION" -foregroundcolor "Green"		
				write-host "***********************************************************************" 	
				write-host "" 						
				write-host "Volume: $Volume" -foregroundcolor "Cyan"
				write-host "Speed: $Speed" -foregroundcolor "Cyan"
				write-host "Voice: $Lang" -foregroundcolor "Cyan"
				write-host "Text: $Text" -foregroundcolor "Cyan"
				write-host "Generate script: $Generate_ps1"	-foregroundcolor "Cyan"
			}
	}								
}