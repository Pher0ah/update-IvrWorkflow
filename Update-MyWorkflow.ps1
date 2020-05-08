#################################################################################################################################################
#RGS Information
#################################################################################################################################################
<#Pool Name      #> $PoolName        = 'tesla-dc-sfbfe.energysafe.vic.gov.au'
<#Name           #> $RGSName         = 'Main IVR'
<#Description    #> $RGSDescription  = 'Main Menu Options'
<#SIP Address    #> $RGSUri          = 'sip:rgs_main@energysafe.vic.gov.au'
<#PSTN Number    #> $RGSLineUri      = 'tel:+61396746395'
<#Display Number #> $RGSNumber       = '03 9674 6395'
<#Business Hours #> $BusinessHours   = 'Always Open'
<#Holiday List   #> $HolidayListName = 'VIC-Holidays'
<#Questions CSV  #> $QuestionsCSV    = 'TestQuestions.csv'

$ServiceID  = "service:ApplicationServer:$($PoolName)"
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
#Create or Update RGS
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

#################################################################################################################################################
#Set Workflow Parameters
#################################################################################################################################################
#Setup Music On Hold
$RGS.CustomMusicOnHoldFile  = $Null

#Set TimeZone and Language Information
$RGS.Language        = 'en-AU'
$RGS.TimeZone        = 'AUS Eastern Standard Time'
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

#################################################################################################################################################
#Update RGS
#################################################################################################################################################
Set-CsRgsWorkflow -Instance $RGS
