<#
.SYNOPSIS
Silly script to invoke random cat facts to the user via speech synth.

.NOTES
Author: John Bignold
#>


function factcat {Add-Type -AssemblyName System.Speech
$SpeechSynth = New-Object System.Speech.Synthesis.SpeechSynthesizer
$SpeechSynth.SelectVoice("Microsoft Zira Desktop")
$Browser = New-Object System.Net.WebClient
$Browser.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$CatFact = (ConvertFrom-Json (Invoke-WebRequest -Verbose -Uri https://catfact.ninja/fact -UseBasicParsing))
$CatFact.fact
$SpeechSynth.Speak("Did you know ?")
$SpeechSynth.Speak($CatFact.fact)}
