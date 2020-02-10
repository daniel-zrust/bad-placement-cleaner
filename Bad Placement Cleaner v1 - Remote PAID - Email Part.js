//Let's send the email
var subjectText = 'Bad Placement Cleaner Just Did Changes in '+accountName+' Account | '+ accountId +' | '+ today

//####
Logger.log('Line milestone email 5')

if(domainEndingsModuleEnabled == 'YES' && domainEndings.length>0 && finalListOfDomainEndingsPlacementsToAdd.length>0){  
  var domainEndingsPlacementsForEmail = 
       'DOMAIN ENDINGS - Here are the bad placements which were added to the "'+domainEndingsListName+'" shared exclusion list:' + '<br>' +
       finalListOfDomainEndingsPlacementsToAdd.join('<br>') +
      '<br>' + '<br>'
  
}

else
  {var domainEndingsPlacementsForEmail = ''}
  

Logger.log('Line milestone email 19')  


//####
if(displayModuleEnabled == 'YES' && typeof finalListOfDisplayPlacementsToAdd !=='undefined' && finalListOfDisplayPlacementsToAdd.length>0){
  var displayPlacementsForEmail =
      'DISPLAY - Here are the bad placements which were added to the "'+displayListName+'" shared exclusion list:' + '<br>' +
      finalListOfDisplayPlacementsToAdd.join('<br>') +
      '<br>' + '<br>'
}
else 
  {var displayPlacementsForEmail = ''}

if(displayModuleEnabled == 'YES' && typeof displayCampaignsWhichGotDomainEndingsExclusionList !== 'undefined' && displayCampaignsWhichGotDomainEndingsExclusionList.length>0){  
  var displayCampaignsforEmailDomainEndingsList =
     'Display campaigns which got Domain Endings exclusion list assigned:'+'<br>'+
      displayCampaignsWhichGotDomainEndingsExclusionList.join('<br>') + 
      '<br>' + '<br>'
}
else 
  {var displayCampaignsforEmailDomainEndingsList = ''}

if(displayModuleEnabled == 'YES' && typeof displayCampaignsWhichGotDisplayExclusionList !== 'undefined' && displayCampaignsWhichGotDisplayExclusionList.length>0){  
  var displayCampaignsforEmailDisplayList =
      'Display campaigns which got Display exclusion list assigned:'+'<br>'+
      displayCampaignsWhichGotDisplayExclusionList.join('<br>') +
      '<br>' + '<br>' 
}
else 
  {var displayCampaignsforEmailDisplayList = ''}

if(displayModuleEnabled == 'YES' && typeof displayCampaignsWhichGotAppsExclusionList !== 'undefined' && displayCampaignsWhichGotAppsExclusionList.length>0){  
  var displayCampaignsforEmailAppList =   
      'Display campaigns which got Apps exclusion list assigned:'+'<br>'+
      displayCampaignsWhichGotAppsExclusionList.join('<br>') +  
      '<br>' + '<br>'
}
else 
    {var displayCampaignsforEmailAppList = ''}

Logger.log('Line milestone email 59') 

//####

  
if(videoModuleEnabled == 'YES' && typeof finalListOfVideoPlacementsToAdd !== 'undefined' && finalListOfVideoPlacementsToAdd.length>0){
  var videoPlacementsForEmail =
      'VIDEO - Here are the bad placements which were added to the "'+videoListName+'" shared exclusion list:' + '<br>' +
      finalListOfVideoPlacementsToAdd.join('<br>') +  
      '<br>' + '<br>'
}
else 
  {var videoPlacementsForEmail = ''}

if(videoModuleEnabled == 'YES' && typeof videoCampaignsWhichGotDomainEndingsExclusionList !== 'undefined' && videoCampaignsWhichGotDomainEndingsExclusionList.length>0){
  var videoCampaignsforEmailDomainEndingsList =       
    'Video campaigns which got Domain Endings exclusion list assigned:'+'<br>'+
    videoCampaignsWhichGotDomainEndingsExclusionList.join('<br>') +     
    '<br>' + '<br>'
}

else 
  {var videoCampaignsforEmailDomainEndingsList = ''}

if(videoModuleEnabled == 'YES' && typeof videoCampaignsWhichGotVideoExclusionList !== 'undefined' && videoCampaignsWhichGotVideoExclusionList.length>0){
  var videoCampaignsforEmailVideoList =
     'Video campaigns which got Video exclusion list assigned:'+'<br>'+
     videoCampaignsWhichGotVideoExclusionList.join('<br>') +     
     '<br>' + '<br>'
}
else 
  {var videoCampaignsforEmailVideoList = ''}


//####
if(appsModuleEnabled == 'YES' && typeof finalListOfAppPlacementsToAdd !== 'undefined' && finalListOfAppPlacementsToAdd.length>0){
  var appsPlacementsForEmail = 'APPS - Here are the bad placements which were added to the "'+appListName+'" shared exclusion list:' + '<br>' +
  finalListOfAppPlacementsToAdd.join('<br>') +
  '<br>' + '<br>'
}

else
  {var appsPlacementsForEmail = ''}   


//Let's contruct the email body
var bodyText =
    ( 'Mode: '+forRealOrTesting + '<br>' + '<br>' +
      
      'License: '+license + '<br>' + '<br>' +

      '***Changes in Bad Placements***:' + '<br>' +
      'Hi,' + '<br>' +

      'Bad Placement Cleaner script has found new bad placements. Here they are: ' + '<br>' + '<br>' +

      domainEndingsPlacementsForEmail +
           
      displayPlacementsForEmail +
      
      videoPlacementsForEmail +
      
      appsPlacementsForEmail +

           
      '***Changes in the Assigned Lists***:' +'<br>'+

      displayCampaignsforEmailDomainEndingsList +
     
      displayCampaignsforEmailDisplayList +
     
      displayCampaignsforEmailAppList +

      videoCampaignsforEmailDomainEndingsList +

      videoCampaignsforEmailVideoList +

      '<br>' + '<br>' +
     
      'Your config spreadsheet is here: https://docs.google.com/spreadsheets/d/'+ssId +
      
      '<br>' + '<br>' + 
      'And that\'s all!' + '<br>' +
      'This script is powered by www.ExcelinPPC.com.' + '<br>' +
      'You can join our newsletter here: https://www.excelinppc.com/newsletter/' + '<br>' +
      'Or you can join our FB group here: https://www.facebook.com/groups/Excelinppc/' + '<br>' +
      'If something is not working, send me an email to mail@danzrust.cz'
    )

if( 
	 email.length>0 && 
    (license == 'paid' || license == 'trial') &&
	
  	(
		    (typeof finalListOfDomainEndingsPlacementsToAdd !== 'undefined' && finalListOfDomainEndingsPlacementsToAdd.length>0 && domainEndingsModuleEnabled == 'YES') || 
		    (typeof finalListOfDisplayPlacementsToAdd !== 'undefined' &&  finalListOfDisplayPlacementsToAdd.lnegth>0 && displayModuleEnabled == 'YES') || 
		    (typeof finalListOfVideoPlacementsToAdd !== 'undefined' && finalListOfVideoPlacementsToAdd.length>0 && videoModuleEnabled == 'YES') || 
		    (typeof finalListOfAppPlacementsToAdd !== 'undefined' && finalListOfAppPlacementsToAdd.length>0 && appsModuleEnabled == 'YES') ||
		    
		    (typeof displayCampaignsWhichGotDomainEndingsExclusionList !== 'undefined' && displayCampaignsWhichGotDomainEndingsExclusionList.length>0 && displayModuleEnabled == 'YES') ||       
		    (typeof displayCampaignsWhichGotDisplayExclusionList !== 'undefined' && displayCampaignsWhichGotDisplayExclusionList.length>0 && displayModuleEnabled == 'YES') || 
		    (typeof displayCampaignsWhichGotAppsExclusionList !== 'undefined' && displayCampaignsWhichGotAppsExclusionList.length>0 && displayModuleEnabled == 'YES') || 

		    (typeof videoCampaignsWhichGotDomainEndingsExclusionList !== 'undefined' && videoCampaignsWhichGotDomainEndingsExclusionList.length>0 && videoModuleEnabled == 'YES') ||
		    (typeof videoCampaignsWhichGotVideoExclusionList !== 'undefined' && videoCampaignsWhichGotVideoExclusionList.length>0 && videoModuleEnabled == 'YES') ||
		    (developerMode == 'developer')
	)
  )

    {
      MailApp.sendEmail({
            to: email, 
            subject: subjectText,
            htmlBody: bodyText
          })

      Logger.log('Email sent.')  
    }

  else 
    {
      Logger.log('Email will not be sent. Couple possible reasons: 1) No new bad placements were added 2) No lists were assigned 3) You are using a free license 4) You did not enter the email address in the config sheet.')
    }  
  