<#
.SYNOPSIS
    Creates or updates a workflow with IVR options specified in this script

.DESCRIPTION
    1. The first section of this script is where you can make changes to put the Workflow parameters and customise the script for your environment.
    2. The second section is where the fun is; you will need to edit this based on the IVR questions and responses; edit this as you see fit; I have tried to include as much as possible in there.
    3. DO NOT EDIT the third section; as this retrieves the Workflow (or creates it); and updates it with the information from the first two sections.

.NOTES
    Version          : 0.01
    Required Infra.  : Lync 2010 or above

    Last Updated     : 15/05/2020
    Wishlist         : Let me know and I will add it
    Known Issues     : No known issues
    Acknowledgements : None

    Author(s)        : Hany Elkady (pheroah@gmail.com)
    Website          : http://pher0ah.blogspot.com.au

    Disclaimer       : The rights to the use of this script is give as is without warranty
                       of any kind implicit or otherwise

.PARAMETER PoolFQDN
    FQDN of the Pool where this IVR will be configured.

.EXAMPLE
    Update-CsIvrWorkflow.ps1 -PoolFQDN sfbFE01.contoso.com.au'

.LINK
    https://pher0ah.blogspot.com/2020/05/complex-ivr-workflow.html

#>
#Requires -Version 2.0
[cmdletbinding()]
param ([Parameter(mandatory=$true)] [string]$PoolFQDN
)

#################################################################################################################################################
#Section 1: RGS Information [Edit this section to fit your environment]
#################################################################################################################################################
<#Name           #> $RGSName         = 'Main IVR'
<#Description    #> $RGSDescription  = 'Main Menu Options'
<#SIP Address    #> $RGSUri          = 'sip:IVR_Main@contoso.com.au'
<#PSTN Number    #> $RGSLineUri      = 'tel:+61398765432'
<#Display Number #> $RGSNumber       = '03 9876 5432'
<#Business Hours #> $BusinessHours   = 'Always Open'
<#Holiday List   #> $HolidayListName = 'VIC-Holidays'

<#IVR Language   #> $IVRLanguage     = 'en-AU'
<#IVR Timezone   #> $IVRTimeZOne     = 'AUS Eastern Standard Time'

$ServiceID  = "service:ApplicationServer:$($PoolFQDN)"
#################################################################################################################################################
#Section 2: IVR Options [Edit this section to customise the workflow IVR Options]
#################################################################################################################################################
$Option0Q = Get-CsRgsQueue -Identity $ServiceID -Owner $PoolFQDN -Name 'Test Queue 0'
$Option1Q = Get-CsRgsQueue -Identity $ServiceID -Owner $PoolFQDN -Name 'Test Queue 1'
$Option2Q = Get-CsRgsQueue -Identity $ServiceID -Owner $PoolFQDN -Name 'Test Queue 2'
$Option3Q = Get-CsRgsQueue -Identity $ServiceID -Owner $PoolFQDN -Name 'Test Queue 3'
$Option4Q = Get-CsRgsQueue -Identity $ServiceID -Owner $PoolFQDN -Name 'Test Queue 4'
$Option5Q = Get-CsRgsQueue -Identity $ServiceID -Owner $PoolFQDN -Name 'Test Queue 5'
$Option6Q = Get-CsRgsQueue -Identity $ServiceID -Owner $PoolFQDN -Name 'Test Queue 6'
$Option7Q = Get-CsRgsQueue -Identity $ServiceID -Owner $PoolFQDN -Name 'Test Queue 7'
$Option8Q = Get-CsRgsQueue -Identity $ServiceID -Owner $PoolFQDN -Name 'Test Queue 8'
$Option9Q = Get-CsRgsQueue -Identity $ServiceID -Owner $PoolFQDN -Name 'Test Queue 9'

#Option 0
$TTS      = 'Blah Blah'
$WavPath  = 'C:\scripts\RGS\Option 0.wav'
$Audio    = Import-CsRgsAudioFile -Identity $ServiceId -FileName "Option 0" -Content (Get-Content -Path $WavPath -Encoding Byte -ReadCount 0)
$Prompt   = New-CsRgsPrompt -TextToSpeechPrompt $TTS -AudioFilePrompt $Audio
$Action   = New-CsRgsCallAction -Prompt $Prompt -Action TransferToQueue -QueueID $Option0Q.Identity
$Answer0  = New-CsRgsAnswer -Name 'Option 0' -Action $Action -DtmfResponse 0 -VoiceResponseList 'Zero', 'Option Zero'

#Option 1
$TTS      = 'Blah Blah'
$WavPath  = 'C:\scripts\RGS\Option 1.wav'
$Audio    = Import-CsRgsAudioFile -Identity $ServiceId -FileName "Option 1" -Content (Get-Content -Path $WavPath -Encoding Byte -ReadCount 0)
$Prompt   = New-CsRgsPrompt -TextToSpeechPrompt $TTS -AudioFilePrompt $Audio
$Action   = New-CsRgsCallAction -Prompt $Prompt -Action TransferToQueue -QueueID $Option1Q.Identity
$Answer1  = New-CsRgsAnswer -Name 'Option 1' -Action $Action -DtmfResponse 1 -VoiceResponseList 'One', 'Option One'

#Option 2
$TTS      = 'Blah Blah'
$WavPath  = 'C:\scripts\RGS\Option 2.wav'
$Audio    = Import-CsRgsAudioFile -Identity $ServiceId -FileName "Option 2" -Content (Get-Content -Path $WavPath -Encoding Byte -ReadCount 0)
$Prompt   = New-CsRgsPrompt -TextToSpeechPrompt $TTS -AudioFilePrompt $Audio
$Action   = New-CsRgsCallAction -Prompt $Prompt -Action TransferToQueue -QueueID $Option2Q.Identity
$Answer2  = New-CsRgsAnswer -Name 'Option 2' -Action $Action -DtmfResponse 2 -VoiceResponseList 'Two', 'Option Two'

#Option 3
$TTS      = 'Blah Blah'
$WavPath  = 'C:\scripts\RGS\Option 3.wav'
$Audio    = Import-CsRgsAudioFile -Identity $ServiceId -FileName "Option 3" -Content (Get-Content -Path $WavPath -Encoding Byte -ReadCount 0)
$Prompt   = New-CsRgsPrompt -TextToSpeechPrompt $TTS -AudioFilePrompt $Audio
$Action   = New-CsRgsCallAction -Prompt $Prompt -Action TransferToQueue -QueueID $Option3Q.Identity
$Answer3  = New-CsRgsAnswer -Name 'Option 3' -Action $Action -DtmfResponse 3 -VoiceResponseList 'Three', 'Option Three'

#Option 4
$TTS      = 'Blah Blah'
$WavPath  = 'C:\scripts\RGS\Option 4.wav'
$Audio    = Import-CsRgsAudioFile -Identity $ServiceId -FileName "Option 4" -Content (Get-Content -Path $WavPath -Encoding Byte -ReadCount 0)
$Prompt   = New-CsRgsPrompt -TextToSpeechPrompt $TTS -AudioFilePrompt $Audio
$Action   = New-CsRgsCallAction -Prompt $Prompt -Action TransferToQueue -QueueID $Option4Q.Identity
$Answer4  = New-CsRgsAnswer -Name 'Option 4' -Action $Action -DtmfResponse 4 -VoiceResponseList 'Four', 'Option Four'

#Option 5
$TTS      = 'Blah Blah'
$WavPath  = 'C:\scripts\RGS\Option 5.wav'
$Audio    = Import-CsRgsAudioFile -Identity $ServiceId -FileName "Option 5" -Content (Get-Content -Path $WavPath -Encoding Byte -ReadCount 0)
$Prompt   = New-CsRgsPrompt -TextToSpeechPrompt $TTS -AudioFilePrompt $Audio
$Action   = New-CsRgsCallAction -Prompt $Prompt -Action TransferToQueue -QueueID $Option5Q.Identity
$Answer5  = New-CsRgsAnswer -Name 'Option 5' -Action $Action -DtmfResponse 5 -VoiceResponseList 'Five', 'Option Five'

#Option 6
$TTS      = 'Blah Blah'
$WavPath  = 'C:\scripts\RGS\Option 6.wav'
$Audio    = Import-CsRgsAudioFile -Identity $ServiceId -FileName "Option 6" -Content (Get-Content -Path $WavPath -Encoding Byte -ReadCount 0)
$Prompt   = New-CsRgsPrompt -TextToSpeechPrompt $TTS -AudioFilePrompt $Audio
$Action   = New-CsRgsCallAction -Prompt $Prompt -Action TransferToQueue -QueueID $Option6Q.Identity
$Answer6  = New-CsRgsAnswer -Name 'Option 6' -Action $Action -DtmfResponse 6 -VoiceResponseList 'Six', 'Option Six'

#Option 7
$TTS      = 'Blah Blah'
$WavPath  = 'C:\scripts\RGS\Option 7.wav'
$Audio    = Import-CsRgsAudioFile -Identity $ServiceId -FileName "Option 7" -Content (Get-Content -Path $WavPath -Encoding Byte -ReadCount 0)
$Prompt   = New-CsRgsPrompt -TextToSpeechPrompt $TTS -AudioFilePrompt $Audio
$Action   = New-CsRgsCallAction -Prompt $Prompt -Action TransferToQueue -QueueID $Option7Q.Identity
$Answer7  = New-CsRgsAnswer -Name 'Option 7' -Action $Action -DtmfResponse 7 -VoiceResponseList 'Seven', 'Option Seven'

#Option 8
$TTS      = 'Blah Blah'
$WavPath  = 'C:\scripts\RGS\Option 8.wav'
$Audio    = Import-CsRgsAudioFile -Identity $ServiceId -FileName "Option 8" -Content (Get-Content -Path $WavPath -Encoding Byte -ReadCount 0)
$Prompt   = New-CsRgsPrompt -TextToSpeechPrompt $TTS -AudioFilePrompt $Audio
$Action   = New-CsRgsCallAction -Prompt $Prompt -Action TransferToQueue -QueueID $Option8Q.Identity
$Answer8  = New-CsRgsAnswer -Name 'Option 8' -Action $Action -DtmfResponse 8 -VoiceResponseList 'Eight', 'Option Eight'

#Option 9
$TTS      = 'Blah Blah'
$WavPath  = 'C:\scripts\RGS\Option 9.wav'
$Audio    = Import-CsRgsAudioFile -Identity $ServiceId -FileName "Option 9" -Content (Get-Content -Path $WavPath -Encoding Byte -ReadCount 0)
$Prompt   = New-CsRgsPrompt -TextToSpeechPrompt $TTS -AudioFilePrompt $Audio
$Action   = New-CsRgsCallAction -Prompt $Prompt -Action TransferToQueue -QueueID $Option9Q.Identity
$Answer9  = New-CsRgsAnswer -Name 'Option 9' -Action $Action -DtmfResponse 9 -VoiceResponseList 'Nine', 'Option Nine'

#Setup Question Prompt
$TTS              = 'Please Press 0 for Option 1, 1 for Option 1, 2 for option 2, ...'
$WavPath          = 'C:\scripts\RGS\IVR_Menu_Options.wav'
$Audio            = Import-CsRgsAudioFile -Identity $ServiceId -FileName "IVR_Menu_Options" -Content (Get-Content -Path $WavPath -Encoding Byte -ReadCount 0)
$IVRPrompt        = New-CsRgsPrompt -TextToSpeechPrompt $TTS -AudioFilePrompt $Audio

#Setup Invalid Prompt
$TTS              = 'Please listen carefully to the options and try again'
$WavPath          = 'C:\cripts\RGS\Invalid_Selection.wav'
$Audio            = Import-CsRgsAudioFile -Identity $ServiceId -FileName "Invalid_Selection" -Content (Get-Content -Path $WavPath -Encoding Byte -ReadCount 0)
$InvalidPrompt    = New-CsRgsPrompt -TextToSpeechPrompt $TTS -AudioFilePrompt $Audio

#Setup No Answer Prompt
$TTS              = 'Are You Still There?'
$WavPath          = 'C:\scripts\RGS\No_Answer.wav'
$Audio            = Import-CsRgsAudioFile -Identity $ServiceId -FileName "No_Answer" -Content (Get-Content -Path $WavPath -Encoding Byte -ReadCount 0)
$NoAnswerPrompt    = New-CsRgsPrompt -TextToSpeechPrompt $TTS -AudioFilePrompt $Audio

#Setup Top Level Question
$TopLevelQuestion = New-CsRgsQuestion -Prompt $IVRPrompt -InvalidAnswerPrompt $InvalidPrompt -NoAnswerPrompt $NoAnswerPrompt `
                                      -AnswerList ($Answer0, $Answer1, $Answer2, $Answer3, $Answer4, $Answer5, $Answer6, $Answer7, $Answer8, $Answer9)

#Setup IVR Action
$TTS        = 'Thank you for calling Blah'
$Prompt     = New-CsRgsPrompt -TextToSpeechPrompt $TTS
$IVRAction  = New-CsRgsCallAction -Prompt $Prompt -Action TransferToQuestion -Question $TopLevelQuestion 

#Setup OOF Action
$TTS       = 'Our working hours are from 8:30 a.m. to 4:30 p.m.'
$WavPath   = 'C:\scripts\RGS\OOF.wav'
$Audio     = Import-CsRgsAudioFile -Identity $ServiceId -FileName "OOF" -Content (Get-Content -Path $WavPath -Encoding Byte -ReadCount 0)
$Prompt    = New-CsRgsPrompt -TextToSpeechPrompt $TTS -AudioFilePrompt $Audio
$OOFAction = New-CsRgsCallAction -Prompt $Prompt -Action Terminate

#Setup Holiday Action
$TTS          = 'We are currently on holiday'
$WavPath   = 'C:\scripts\RGS\Holiday.wav'
$Audio     = Import-CsRgsAudioFile -Identity $ServiceId -FileName "Holiday" -Content (Get-Content -Path $WavPath -Encoding Byte -ReadCount 0)
$Prompt    = New-CsRgsPrompt -TextToSpeechPrompt $TTS -AudioFilePrompt $Audio
$HolidayAction = New-CsRgsCallAction -Prompt $Prompt -Action Terminate

#################################################################################################################################################
#Section 3: Create or Update RGS [DO NOT EDIT BELOW THIS LINE]
#################################################################################################################################################
Remove-Variable -Name 'RGS' -Force -ErrorAction SilentlyContinue
$RGS        = (Get-CsRgsWorkflow -Identity $ServiceID -Owner $PoolFQDN -Name $RGSName -ErrorAction SilentlyContinue)

If($RGS){
  #Update Workflow Parameters
  $RGS.Name                 = $RGSName
  $RGS.Description          = $RGSDescription
  $RGS.DisplayNumber        = $RGSNumber
  $RGS.LineURI              = $RGSLineUri
  $RGS.Active               = $true
  $RGS.Anonymous            = $false
  $RGS.EnabledForFederation = $false
  $RGS.Managed              = $false

}else{
  #Create a new RGS
  $NewRGS = '$RGS = '
  $NewRGS += "New-CsRgsWorkflow -Parent '$($ServiceID)' -Name '$($RGSName)' -Description '$($RGSDescription)' -PrimaryUri '$($RGSUri)'"
  $NewRGS += ' -Active $true -Anonymous $false -EnabledForFederation $false -Managed $false -DefaultAction $IVRAction'
  If($RGSNumber){$NewRGS += " -DisplayNumber '$($RGSNumber)' -LineUri '$($RGSLineUri)'"}
  Invoke-Expression -Command $NewRGS -OutVariable $RGS                                             
}

########################
#Set Workflow Parameters
########################

#Setup Music On Hold
$RGS.CustomMusicOnHoldFile  = $Null

#Set TimeZone and Language Information
$RGS.Language        = $IVRLanguage
$RGS.TimeZone        = $IVRTimeZone
$RGS.BusinessHoursID = (Get-CsRgsHoursOfBusiness -Name $BusinessHours).Identity[0]

#Set Managment
$RGS.Managed                = $False
$RGS.ManagersByUri.Clear()

#Set Holiday List
$RGS.HolidaySetIDList.Clear()
If($HolidaylistName){
  $RGS.HolidaySetIDList.Add((Get-CsRgsHolidaySet -Name $HolidayListName).Identity)
}

#Update Actions
$RGS.NonBusinessHoursAction = $OOFAction
$RGS.HolidayAction          = $HolidayAction
$RGS.DefaultAction          = $IVRAction

#Update RGS
Set-CsRgsWorkflow -Instance $RGS
