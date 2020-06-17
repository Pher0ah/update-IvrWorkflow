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
<#Pool Name      #> $PoolName        = 'sfbFE01.contoso.com.au' <#--------------- TESTING ONLY #>
<#Name           #> $RGSName         = 'Main IVR'
<#Description    #> $RGSDescription  = 'Main Menu Options'
<#SIP Address    #> $RGSUri          = 'sip:IVR_Main@contoso.com.au'
<#PSTN Number    #> $RGSLineUri      = 'tel:+61398765432'
<#Display Number #> $RGSNumber       = '03 9876 5432'
<#Business Hours #> $BusinessHours   = 'Always Open'
<#Holiday List   #> $HolidayListName = 'VIC-Holidays'
<#Questions CSV  #> $QuestionsCSV    = 'TestQuestions.csv'

<#IVR Language   #> $IVRLanguage     = 'en-AU'
<#IVR Timezone   #> $IVRTimeZOne     = 'AUS Eastern Standard Time'

$ServiceID  = "service:ApplicationServer:$($PoolName)"
#################################################################################################################################################
#Section 2: IVR Options [Edit this section to customise the workflow IVR Options]
#################################################################################################################################################


#################################################################################################################################################
# Read Questions File
#################################################################################################################################################
$QuestionsInfo = Import-Csv -Delimiter ',' -Path $QuestionsCSV

If($QuestionsInfo.count -eq 0){
  write-error "ERROR: Questions File is Empty"
  Exit
}

#################################################################################################################################################
# Process Question Options
#################################################################################################################################################
$Answers =@()

foreach ($QuestionInfo in $QuestionsInfo){

  #Skip Non-Numeric
  If($QuestionInfo.DTMF -NotMatch "[\d\.]+$"){continue}

  #Get Queue
  $Queue = Get-CsRgsQueue -Identity $ServiceID -Owner $PoolName -Name $QuestionInfo.Queue

  #Create Prompt
  $OptionName = "Option-$($QuestionInfo.DTMF)"
  $Audio = Import-CsRgsAudioFile -Identity $ServiceId `
                                 -FileName $OptionName `
                                 -Content (Get-Content -Path $QuestionInfo.Audio -Encoding Byte -ReadCount 0)
  $Prompt= New-CsRgsPrompt -TextToSpeechPrompt $QuestionInfo.TTS -AudioFilePrompt $Audio

  #Create Action
  $Action   = New-CsRgsCallAction -Prompt $Prompt -Action TransferToQueue -QueueID $Queue.Identity

  #Add Answer to List
  $Answers += New-CsRgsAnswer -Name $OptionName -Action $Action -DtmfResponse $QuestionInfo.DTMF -VoiceResponseList ($QuestionInfo.Voice1, $QuestionInfo.Voice2)
}


#################################################################################################################################################
# Setup Top Level Question
#################################################################################################################################################
$Audio            = Import-CsRgsAudioFile -Identity $ServiceId `
                                          -FileName "Question-Prompt" `
                                          -Content (Get-Content -Path $QuestionsInfo[0].Audio -Encoding Byte -ReadCount 0)
$QuestionPrompt   = New-CsRgsPrompt -TextToSpeechPrompt $QuestionsInfo[0].TTS -AudioFilePrompt $Audio

$Audio            = Import-CsRgsAudioFile -Identity $ServiceId `
                                          -FileName "Response-Invalid" `
                                          -Content (Get-Content -Path $QuestionsInfo[1].Audio -Encoding Byte -ReadCount 0)
$InvalidPrompt    = New-CsRgsPrompt -TextToSpeechPrompt $QuestionsInfo[1].TTS -AudioFilePrompt $Audio

$Audio            = Import-CsRgsAudioFile -Identity $ServiceId `
                                          -FileName "No-Response" `
                                          -Content (Get-Content -Path $QuestionsInfo[2].Audio -Encoding Byte -ReadCount 0)
$NoAnswerPrompt   = New-CsRgsPrompt -TextToSpeechPrompt $QuestionsInfo[2].TTS -AudioFilePrompt $Audio

$TopLevelQuestion = New-CsRgsQuestion -Prompt $IVRPrompt -InvalidAnswerPrompt $InvalidPrompt -NoAnswerPrompt $NoAnswerPrompt -AnswerList $Answers

#Setup Default Action
$TTS        = 'Thank you for calling Blah'
$Prompt     = New-CsRgsPrompt -TextToSpeechPrompt $TTS
$DefaultAction  = New-CsRgsCallAction -Prompt $Prompt -Action TransferToQuestion -Question $TopLevelQuestion 

#Setup OOF Action
$TTS       =  'Our working hours are from 8:30 a.m. to 4:30 p.m.'
$Prompt    = New-CsRgsPrompt -TextToSpeechPrompt $TTS
$OOFAction = New-CsRgsCallAction -Prompt $Prompt -Action Terminate

#Setup Holiday Action
$TTS          = 'We are currently on holiday'
$Prompt        = New-CsRgsPrompt -TextToSpeechPrompt $TTS
$HolidayAction = New-CsRgsCallAction -Prompt $Prompt -Action Terminate

#################################################################################################################################################
#Section 3: Create or Update RGS [DO NOT EDIT BELOW THIS LINE]
#################################################################################################################################################
Remove-Variable -Name 'RGS' -Force -ErrorAction SilentlyContinue
$RGS        = (Get-CsRgsWorkflow -Identity $ServiceID -Owner $PoolName -Name $RGSName -ErrorAction SilentlyContinue)

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
  $NewRGS += ' -Active $true -Anonymous $false -EnabledForFederation $false -Managed $false -DefaultAction $DefaultAction'
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
