//Copy and paste this into JS!!
  //***************************DO NOT TOUCH ANYTHING BELOW*****************
  //send event started to GA if allowed
  //Let's start doing the magic
  
  var ss = SpreadsheetApp.openById(ssId)
  //set locale to EN-US so number formats are suitable for all the query functions.
  
  ss.setSpreadsheetLocale('en-US')
  var eventReportingToExcelinPpcComAllowed = ss.getRangeByName('reporting_allowed').getValue()
  var payload = {
    'v': '1',
    'tid': 'UA-69459605-1',
    't': 'event',
    'cid' : Utilities.getUuid(),
    'z' : Math.floor(Math.random() * 10E7),
    'ds' : 'google_ads_script',
    'cn' : 'script-bad_placement_cleaner',
    'cs' : 'google_ads_script',
    'cm' : 'google_ads_script',
    'ec' : 'script_run',
    'ea' : 'script_run_started',
    'el' : 'script-bad_placement_cleaner'
    };
  
  var options = {
    "method" : "post",
    "payload" : payload
   };
  
  //Send the hit to GA if allowed
  if(eventReportingToExcelinPpcComAllowed == 'YES'){UrlFetchApp.fetch("http://www.google-analytics.com/collect", options)};  

  //Date
  var now = new Date()
  var today = Utilities.formatDate(new Date(now.getTime() - 1000 * 60 * 60 * 24), "GMT+1", "yyyy-MM-dd")
        
  //Licensing magic
  var accountName = AdsApp.currentAccount().getName()

  //licenseId
  if(ss.getRangeByName('license_id').getValue().length==0){
    var licenseId = Utilities.getUuid()
    ss.getRangeByName('license_id').setValue(licenseId)
  }
  else
    {var licenseId = ss.getRangeByName('license_id').getValue()}
  
  //accountId
  var accountId = AdsApp.currentAccount().getCustomerId()
   
  if(typeof scriptVersion !== 'undefined')
    {var scriptVersionFinal = scriptVersion} 
    
    else {var scriptVersionFinal = 'scriptVersion variable not present'}      


  var userStatusRequest = UrlFetchApp.fetch('https://hook.integromat.com/xl5vz6d6l5u1yycv7styjsnig7fxha1e?licenseId='+licenseId+'&accountId='+accountId+'&scriptVersion='+scriptVersionFinal).getContentText()
  var license = userStatusRequest

  if(license == 'paid'){Logger.log('Great! You have a paid license so you will get an email in case changes are made.')}
  else if(license == 'trial'){Logger.log('Great! You have a trial license so you will get an email in case changes are made.')}
  else if(license == 'free'){Logger.log('Sorry, you have a free license - no email for you.')}
  else {Logger.log('Sorry, the script failed to verify your license status. Try to re-run it. The script will likely fail in one of the next steps.')}
  
  //Basic variables
  var forRealOrTesting = ss.getRangeByName('for_real_or_testing').getValue()
  var timeRange = ss.getRangeByName('time_range').getValue()
  var email = ss.getRangeByName('email').getValue()
  var statsFromEnabledCampaigns = ss.getRangeByName('stats_from_enabled_campaigns').getValue()
  var minimumImpressions =  ss.getRangeByName('minimum_impressions').getValue()
  var loggerSetup =  ss.getRangeByName('logger_setup').getValue()
  var developerMode = ss.getRangeByName('developer_mode').getValue()

  var domainEndingsListName = ss.getRangeByName('domain_endings_negative_placement_list_name').getValue()  
  var displayListName = ss.getRangeByName('display_negative_placement_list_name').getValue()  
  var videoListName = ss.getRangeByName('video_negative_placement_list_name').getValue()   
  var appListName = ss.getRangeByName('app_negative_placement_list_name').getValue()   

  var domainEndings = ss.getRangeByName('domain_endings').getValue()
    
  var displayQueryAggregation = ss.getRangeByName('display_query_aggregation').getValue()
  var displayQueryFiltering = ss.getRangeByName('display_query_filtering').getValue()
  
  var videoQueryAggregation = ss.getRangeByName('video_query_aggregation').getValue()
  var videoQueryFiltering = ss.getRangeByName('video_query_filtering').getValue()  
  var videoBumperExclusion = ss.getRangeByName('video_bumper_exclusion').getValue()
  var videoBumperString = ss.getRangeByName('video_bumper_string').getValue()
  
  var appsQueryAggregation = ss.getRangeByName('apps_query_aggregation').getValue()
  var appsQueryFiltering = ss.getRangeByName('apps_query_filtering').getValue() 
  var appsMobileUrlFormula = ss.getRangeByName('apps_mobile_urls_formula').getValue()
  
  //Let's define sheet names
  var sheetNameCampaignList = 'script_campaign_list_from_google_ads'
  
  var sheetNameDomainEndingsExport = 'script_domain_endings_bad_placements'

  var sheetNameReportExportDisplay = 'script_display_export_from_gads'
  var sheetNameReportExportVideo = 'script_video_export_from_gads'
  var sheetNameReportExportApps = 'script_apps_export_from_gads'
    
  var displayAggregatedPlacementData = 'script_display_agg_placement_data'
  var displayFilteredBadPlacements = 'script_display_filtered_bad_placements'
  var displayListOfDisplayCampaigns = 'script_display_list_of_display_campaigns'
  
  var videoAggregatedPlacementData = 'script_video_agg_placement_data'
  var videoFilteredBadPlacements = 'script_video_filtered_bad_placements'

  var appsAggregatedPlacementData = 'script_apps_agg_placement_data'
  var appsFilteredBadPlacements = 'script_apps_filtered_bad_placements'
  
  //Module status
  var domainEndingsModuleEnabled =  ss.getRangeByName('domain_endings_enabled').getValue()
  var displayModuleEnabled =  ss.getRangeByName('display_module_enabled').getValue()
  var displayAnonymousPlacementEnabled =  ss.getRangeByName('display_anonymous_yes_or_no').getValue()
  
  var videoModuleEnabled =  ss.getRangeByName('video_module_enabled').getValue()
  var appsModuleEnabled =  ss.getRangeByName('apps_module_enabled').getValue()
  
  //Let's get list of named ranges
  var namedRanges = ss.getNamedRanges()
  var listOfNamedRanges = []
  
  for(i=0;i<namedRanges.length;i++){
    var namedRange = namedRanges[i].getName()
    listOfNamedRanges.push(namedRange)
  }
    
  if(loggerSetup == 'HEAVY'){Logger.log('List of named ranges: '+listOfNamedRanges)}  
  
  //Let's prepare our spreadsheet
  //Let's delete the helper sheets from the previous runs by deleting all sheets which start with 'script_'  
  var sheets = ss.getSheets()
  var sheetNameStartsWith = new RegExp('^script_')
  var sheetNames = []

  for(i=0;i<sheets.length;i++){
    var sheetName = sheets[i].getName()  
    
    if(sheetName.toString().match(sheetNameStartsWith)!==null){
      sheetNames.push(sheetName)
    }
  }

  Logger.log('List of existing helper sheets: '+sheetNames)
    
  //Then delete them
  for(i=0;i<sheetNames.length;i++){
    ss.deleteSheet(ss.getSheetByName(sheetNames[i]))
    } 

  ss.setActiveSheet(ss.getSheetByName('formula_config_do_not_touch'))
  ss.moveActiveSheet(2)
                                                                          
  //Let's create the sheets again
  ss.insertSheet(sheetNameCampaignList)
  var sheetNameCampaignListSheet = ss.getSheetByName(sheetNameCampaignList)
  sheetNameCampaignListSheet.deleteColumns(1,25)   
  sheetNameCampaignListSheet.setTabColor('yellow')

  ss.insertSheet(sheetNameDomainEndingsExport)
  var sheetNameDomainEndingsExportSheet = ss.getSheetByName(sheetNameDomainEndingsExport)
  sheetNameDomainEndingsExportSheet.deleteColumns(1,25)   
  sheetNameDomainEndingsExportSheet.setTabColor('blue')
  sheetNameDomainEndingsExportSheet.deleteRows(1,999)
  sheetNameDomainEndingsExportSheet.getRange('A1').setValue('Placements to exclude based on domain endings')
  
  ss.insertSheet(sheetNameReportExportDisplay)
  var sheetNameReportExportSheetDisplay = ss.getSheetByName(sheetNameReportExportDisplay)
  sheetNameReportExportSheetDisplay.deleteColumns(1,25) 
  sheetNameReportExportSheetDisplay.setTabColor('black')
  
  ss.insertSheet(displayAggregatedPlacementData)
  var displayAggregatedPlacementDataSheet = ss.getSheetByName(displayAggregatedPlacementData)
  displayAggregatedPlacementDataSheet.deleteColumns(13,13)
  displayAggregatedPlacementDataSheet.setTabColor('black')
  displayAggregatedPlacementDataSheet.deleteRows(1,999)
                                                                          
  ss.insertSheet(displayFilteredBadPlacements)
  var displayFilteredBadPlacementsSheet = ss.getSheetByName(displayFilteredBadPlacements)
  displayFilteredBadPlacementsSheet.deleteColumns(13,13)
  displayFilteredBadPlacementsSheet.setTabColor('black')
  displayFilteredBadPlacementsSheet.deleteRows(1,990)
  
  ss.insertSheet(sheetNameReportExportVideo)
  var sheetNameReportExportSheetVideo = ss.getSheetByName(sheetNameReportExportVideo)
  sheetNameReportExportSheetVideo.deleteColumns(1,25) 
  sheetNameReportExportSheetVideo.setTabColor('orange')  
  
  ss.insertSheet(videoAggregatedPlacementData)
  var videoAggregatedPlacementDataSheet = ss.getSheetByName(videoAggregatedPlacementData)
  videoAggregatedPlacementDataSheet.deleteColumns(13,13)  
  videoAggregatedPlacementDataSheet.setTabColor('orange')
  videoAggregatedPlacementDataSheet.deleteRows(1,990)
  
  ss.insertSheet(videoFilteredBadPlacements)
  var videoFilteredBadPlacementsSheet = ss.getSheetByName(videoFilteredBadPlacements)
  videoFilteredBadPlacementsSheet.deleteColumns(13,13)
  videoFilteredBadPlacementsSheet.setTabColor('orange')
  videoFilteredBadPlacementsSheet.deleteRows(1,990)
  
  ss.insertSheet(sheetNameReportExportApps)
  var sheetNameReportExportSheetApps = ss.getSheetByName(sheetNameReportExportApps)
  sheetNameReportExportSheetApps.deleteColumns(1,25)   
  sheetNameReportExportSheetApps.setTabColor('purple')    
  
  ss.insertSheet(appsAggregatedPlacementData)
  var appsAggregatedPlacementDataSheet = ss.getSheetByName(appsAggregatedPlacementData)
  appsAggregatedPlacementDataSheet.deleteColumns(13,13)
  appsAggregatedPlacementDataSheet.setTabColor('purple')
  appsAggregatedPlacementDataSheet.deleteRows(1,990)

  ss.insertSheet(appsFilteredBadPlacements)
  var appsFilteredBadPlacementsSheet = ss.getSheetByName(appsFilteredBadPlacements)
  appsFilteredBadPlacementsSheet.deleteColumns(14,12)  
  appsFilteredBadPlacementsSheet.setTabColor('purple')
  appsFilteredBadPlacementsSheet.deleteRows(1,990)
  
  //Let's sort out domain endings problem first
  if(domainEndings.length>0 && domainEndingsModuleEnabled == 'YES'){
    Logger.log('')
    Logger.log('***Running Domain Endings module***')
    Logger.log('')
    
    var domainEndingsAsList = domainEndings.split('|')
    var domainEndingsAsListWithRegex = []

    for(i=0;i<domainEndingsAsList.length;i++){
      var endingWithRegex = domainEndingsAsList[i].trim().concat('$')
      domainEndingsAsListWithRegex.push(endingWithRegex)
    }

    var domainEndingsAsStringtWithRegex = domainEndingsAsListWithRegex.join('|')
    Logger.log('Domain endings with regex: '+domainEndingsAsStringtWithRegex)  
  

    var placementsToCheck = []
    var placementsToBeExcludedBasedOnEndings = []
    
    var queryDomainEndings =  'SELECT Criteria '+ 
                              'FROM PLACEMENT_PERFORMANCE_REPORT ' +
                              'WHERE AdNetworkType1 IN [CONTENT, YOUTUBE_SEARCH, YOUTUBE_WATCH, MIXED] AND Impressions > 0 ' +
                              'DURING TODAY' 

    Logger.log('Running report for domain endings check.')
    var domainEndingsReport = AdsApp.report(queryDomainEndings)
    
    var rows = domainEndingsReport.rows()
    
    while(rows.hasNext()){
      var eachRow = rows.next()
      var placement = eachRow['Criteria'] 

      placementsToCheck.push(placement)
      
    }
        
    reg = new RegExp(domainEndingsAsStringtWithRegex+'$')
    
    for(i=0;i<placementsToCheck.length;i++){
      //This is the check for "domain ends with"
      if(placementsToCheck[i].toString().match(reg)!==null){
        placementsToBeExcludedBasedOnEndings.push(placementsToCheck[i])
        if(loggerSetup == 'HEAVY'){Logger.log(placementsToCheck[i] + ' added to excluded placements - domain ending matched.')}
      }
    }
  
  var placementsToBeExcludedBasedOnEndingsUnique = placementsToBeExcludedBasedOnEndings.sort().filter(function(item, index){
      return placementsToBeExcludedBasedOnEndings.indexOf(item) >= index;
    })  
    
  var countOfPlacementsToBeExcludeDomainEndings = placementsToBeExcludedBasedOnEndingsUnique.length  
  Logger.log('countOfPlacementsToExcludeDomainEndings: '+ placementsToBeExcludedBasedOnEndingsUnique.length)
  Logger.log('These are the placements to exclude based on domain endings: '+ placementsToBeExcludedBasedOnEndingsUnique)  
  Logger.log('')
  Logger.log('***I\'m done with identification of bad placements in Domain Endings module***')
  Logger.log('')    
    
  }


  //Let's get campaign lists: Display + Video
  //For getting correct classification, I have to look into AWQL results because you cannot get just display campaigns via campaign selector in scripts.
  //So first, I need to run campaign performance report
  var queryCampaignList =
       'SELECT AccountDescriptiveName, AdvertisingChannelType, AdNetworkType1, AdNetworkType2, CampaignName, CampaignStatus, Impressions, ActiveViewImpressions, VideoViews, Clicks, Cost, ViewThroughConversions, Conversions '+ 
       'FROM CAMPAIGN_PERFORMANCE_REPORT ' +
       'WHERE Impressions > 0'
 
  Logger.log(queryCampaignList)
  var report = AdsApp.report(queryCampaignList)
  report.exportToSheet(sheetNameCampaignListSheet)    
  
  //Than I can start creating Video and Display campaign list
  var campaignNames = sheetNameCampaignListSheet.getRange('B2:E').getValues()
  Logger.log('Campaign names: '+campaignNames)
  
  var lastRowCampaignNames = sheetNameCampaignListSheet.getLastRow()-1 //minus one because the first row is the header
  Logger.log('Count of campaign names: '+lastRowCampaignNames)
  
  var displayCampaignNames = []
 
  for(z=0;z<lastRowCampaignNames;z++){
    var advertisingChannelType = campaignNames[z][0]
    //Logger.log(campaignNames[z][0])
    var campaignName = campaignNames[z][3]
    
    if(advertisingChannelType == 'Display'){displayCampaignNames.push(campaignName)}
  }
    
  Logger.log('List of Display campaigns: '+ displayCampaignNames)
  
  var videoCampaignNames = []
  
  for(t=0;t<lastRowCampaignNames;t++){
    var advertisingChannelType = campaignNames[t][0]
    var campaignName = campaignNames[t][3]
    
    if(advertisingChannelType == 'Video'){videoCampaignNames.push(campaignName)}
  }
    
  Logger.log('List of Video campaigns: '+ videoCampaignNames)  
    
  //Let's run the AWQL reports for display/video/apps only if any of the module is enabled
  if(displayModuleEnabled == 'YES' || videoModuleEnabled == 'YES' || appsModuleEnabled == 'YES'){   

    if(statsFromEnabledCampaigns=='Get stats only from enabled campaigns')
      {var whereCondition = "CampaignName IN ['"+displayCampaignNames.join('\',\'')+"'] AND CampaignStatus = ENABLED AND Impressions >= " + minimumImpressions} 
       
    else
      {var whereCondition = "CampaignName IN ['"+displayCampaignNames.join('\',\'')+"'] AND Impressions >= " + minimumImpressions} 
    
    if(timeRange == 'ALL_TIME')
      {
      var queryDisplay =
          'SELECT AccountDescriptiveName, AdNetworkType1, CampaignName, AdGroupName, Criteria, Impressions, ActiveViewImpressions, VideoViews, Clicks, Cost, ViewThroughConversions, Conversions '+ 
          'FROM PLACEMENT_PERFORMANCE_REPORT ' +
          'WHERE AdNetworkType1 IN [CONTENT, YOUTUBE_SEARCH, YOUTUBE_WATCH, MIXED] AND ' + whereCondition
      }
    
    else
      {
      var queryDisplay =
          'SELECT AccountDescriptiveName, AdNetworkType1, CampaignName, AdGroupName, Criteria, Impressions, ActiveViewImpressions, VideoViews, Clicks, Cost, ViewThroughConversions, Conversions '+ 
          'FROM PLACEMENT_PERFORMANCE_REPORT ' +
          'WHERE AdNetworkType1 IN [CONTENT, YOUTUBE_SEARCH, YOUTUBE_WATCH, MIXED] AND ' + whereCondition + ' ' +
          'DURING '+timeRange
      } 
    
    Logger.log('Display query: '+queryDisplay)
    if(displayCampaignNames.length>0){  
      var reportDisplay = AdsApp.report(queryDisplay)
      reportDisplay.exportToSheet(sheetNameReportExportSheetDisplay)
    }
    else 
      {Logger.log('No display campaigns in the account - display query will not run.')}  
    
    //Then we can run placement report for Video campaigns
    //In this case, we need to take care of bumper campaigns unlike for the Display campaigns above
    Logger.log('Stats setup: '+statsFromEnabledCampaigns +' | Bumper exclusion: '+videoBumperExclusion+'| String: '+videoBumperString + ' | Bumper string length: '+videoBumperString.trim.length)
    
    if(statsFromEnabledCampaigns=='Get stats only from enabled campaigns' && videoBumperExclusion == 'YES' && videoBumperString.length>0)
      {var whereCondition = "CampaignName IN ['"+videoCampaignNames.join('\',\'')+"'] AND CampaignName DOES_NOT_CONTAIN '"+videoBumperString+"' AND CampaignStatus = ENABLED AND Impressions >= " + minimumImpressions}
    
    else if(statsFromEnabledCampaigns=='Get stats even from paused campaigns' && videoBumperExclusion == 'YES' && videoBumperString.length>0)
      {var whereCondition = "CampaignName IN ['"+videoCampaignNames.join('\',\'')+"'] AND CampaignName DOES_NOT_CONTAIN '"+videoBumperString+"' AND Impressions >= " + minimumImpressions}
    
    else if(statsFromEnabledCampaigns=='Get stats only from enabled campaigns')
      {var whereCondition = "CampaignName IN ['"+videoCampaignNames.join('\',\'')+"'] AND CampaignStatus = ENABLED AND Impressions >= " + minimumImpressions} 
       
    else
      {var whereCondition = "CampaignName IN ['"+videoCampaignNames.join('\',\'')+"'] AND Impressions >= " + minimumImpressions} 
    
    if(timeRange == 'ALL_TIME')
      {
      var queryVideo =
          'SELECT AccountDescriptiveName, AdNetworkType1, CampaignName, AdGroupName, Criteria, Impressions, Clicks, Cost, Conversions, ViewThroughConversions, VideoViews, VideoViewRate '+ 
          'FROM PLACEMENT_PERFORMANCE_REPORT ' +
          'WHERE AdNetworkType1 IN [CONTENT, YOUTUBE_SEARCH, YOUTUBE_WATCH, MIXED] AND ' + whereCondition
      }
    
    else
      {
      var queryVideo =
          'SELECT AccountDescriptiveName, AdNetworkType1, CampaignName, AdGroupName, Criteria, Impressions, Clicks, Cost, Conversions, ViewThroughConversions, VideoViews, VideoViewRate '+ 
          'FROM PLACEMENT_PERFORMANCE_REPORT ' +
          'WHERE AdNetworkType1 IN [CONTENT, YOUTUBE_SEARCH, YOUTUBE_WATCH, MIXED] AND ' + whereCondition + ' ' +
          'DURING '+timeRange
      } 
    
    Logger.log('Video query: '+queryVideo)
    if(videoCampaignNames.length>0){  
      var reportVideo = AdsApp.report(queryVideo)
      reportVideo.exportToSheet(sheetNameReportExportSheetVideo)
    }
    else 
      {Logger.log('No video campaigns in the account - video query will not run.')}
  
  }

  //Let's sync the sheet with all the changes
  SpreadsheetApp.flush() 

  Logger.log('Line milestone 401')
  
  //Let's add the query functions  
  //But we needs to prepare the apps data first
  //Let's combine the Display + Video data first so we get full list of apps
  // roAdjustment is there to make sure the methods do not get outside of sheet dimensions.
  var rowAdjustment = 3
  
  //I need to check both for Display and Apps enabled since Display is providing data to Apps module
  if(displayModuleEnabled == 'YES' || appsModuleEnabled == 'YES'){
    var lastRowDisplayPerformance = ss.getSheetByName(sheetNameReportExportDisplay).getLastRow()+rowAdjustment
    Logger.log('Count of Display performance rows: '+lastRowDisplayPerformance)
    }

  if(videoModuleEnabled == 'YES' || appsModuleEnabled == 'YES'){  
    var lastRowVideoPerformance =  ss.getSheetByName(sheetNameReportExportVideo).getLastRow()+rowAdjustment
    Logger.log('Count of Video performance rows: '+lastRowVideoPerformance)
  }
  
  if(domainEndings.length>0 && domainEndingsModuleEnabled == 'YES'){
    //Let's add the needed number of rows for domain endings sheet
    sheetNameDomainEndingsExportSheet.insertRowsAfter(1,countOfPlacementsToBeExcludeDomainEndings+rowAdjustment)
  }
  
  //Let's add the needed number of rows for each agg and filter sheet
  if(displayModuleEnabled == 'YES' || appsModuleEnabled == 'YES'){  
    displayAggregatedPlacementDataSheet.insertRowsAfter(1,(lastRowDisplayPerformance))
    displayFilteredBadPlacementsSheet.insertRowsAfter(1,(lastRowDisplayPerformance))
  }

  Logger.log('Line milestone 431')

  if(videoModuleEnabled == 'YES' || appsModuleEnabled == 'YES'){
    videoAggregatedPlacementDataSheet.insertRowsAfter(1,(lastRowVideoPerformance))
    videoFilteredBadPlacementsSheet.insertRowsAfter(1,(lastRowVideoPerformance))
  }

  if(appsModuleEnabled == 'YES'){
    appsAggregatedPlacementDataSheet.insertRowsAfter(1,(lastRowDisplayPerformance+lastRowVideoPerformance))
    appsFilteredBadPlacementsSheet.insertRowsAfter(1,(lastRowDisplayPerformance+lastRowVideoPerformance))
  
    ss.getSheetByName(sheetNameReportExportApps).getRange('A1').setValue('AccountDescriptiveName')
    ss.getSheetByName(sheetNameReportExportApps).getRange('B1').setValue('AdNetworkType1')
    ss.getSheetByName(sheetNameReportExportApps).getRange('C1').setValue('CampaignName')
    ss.getSheetByName(sheetNameReportExportApps).getRange('D1').setValue('AdGroupName')
    ss.getSheetByName(sheetNameReportExportApps).getRange('E1').setValue('Criteria')
    ss.getSheetByName(sheetNameReportExportApps).getRange('F1').setValue('Impressions')
    ss.getSheetByName(sheetNameReportExportApps).getRange('G1').setValue('ActiveViewImpressions')
    ss.getSheetByName(sheetNameReportExportApps).getRange('H1').setValue('VideoViews')
    ss.getSheetByName(sheetNameReportExportApps).getRange('I1').setValue('Clicks')
    ss.getSheetByName(sheetNameReportExportApps).getRange('J1').setValue('Cost')
    ss.getSheetByName(sheetNameReportExportApps).getRange('K1').setValue('ViewThroughConversions')
    ss.getSheetByName(sheetNameReportExportApps).getRange('L1').setValue('Conversions')  
  
    
    SpreadsheetApp.flush()

    Logger.log('Line milestone 456')

    //now copy the Display data over to the apps sheet
    try{ss.getSheetByName(sheetNameReportExportDisplay).getRange('A2:'+lastRowDisplayPerformance).copyTo(sheetNameReportExportSheetApps.getRange('A2'), {contentsOnly:true})} catch(err){Logger.log('No display placements - nothing to copy to '+sheetNameReportExportApps+' sheet')}

    //paste Video data below Display data...
    try{
        sheetNameReportExportSheetApps.insertRows(lastRowDisplayPerformance-rowAdjustment, 4)
        ss.getSheetByName(sheetNameReportExportVideo).getRange('A2:'+lastRowVideoPerformance).copyTo(sheetNameReportExportSheetApps.getRange('A'+(lastRowDisplayPerformance)), {contentsOnly:true})} 

    catch(err){Logger.log('No video placements - nothing to copy to '+sheetNameReportExportApps+' sheet')}  
  }
  
  //Let's dump the placements from domain endins module onto the sheet
  if(domainEndings.length>0 && domainEndingsModuleEnabled == 'YES' && countOfPlacementsToBeExcludeDomainEndings>0){
    SpreadsheetApp.flush()
    
    var placementsToExcludeBasedOnEndingsAsArrayOfArrays = []
    
    for(i=0;i<placementsToBeExcludedBasedOnEndingsUnique.length;i++){
      var eachUrlAsArray = placementsToBeExcludedBasedOnEndingsUnique[i].split()
      placementsToExcludeBasedOnEndingsAsArrayOfArrays.push(eachUrlAsArray)
      }      
    
    sheetNameDomainEndingsExportSheet.getRange('A2:A'+(countOfPlacementsToBeExcludeDomainEndings+1)).setValues(placementsToExcludeBasedOnEndingsAsArrayOfArrays)
  }
  
  Logger.log('Line milestone 481')

  //now add the query functions
  if(displayModuleEnabled == 'YES' || appsModuleEnabled == 'YES'){    
    var displayAggregationFormula = '='+displayQueryAggregation.replace(/script_display_export_from_gads/g,sheetNameReportExportDisplay)                                                                                 
    displayAggregatedPlacementDataSheet.getRange('A1').setFormula(displayAggregationFormula)
  
    var displayFilterFormula = '='+displayQueryFiltering.replace(/script_display_agg_placement_data/g,displayAggregatedPlacementData)
    displayFilteredBadPlacementsSheet.getRange('A1').setFormula(displayFilterFormula)     
  }

  if(videoModuleEnabled == 'YES' || appsModuleEnabled == 'YES'){
    var videoAggregationFormula = '='+videoQueryAggregation.replace(/script_video_export_from_gads/g,sheetNameReportExportVideo)                                                                                 
    videoAggregatedPlacementDataSheet.getRange('A1').setFormula(videoAggregationFormula) 

    var videoFilterFormula = '='+videoQueryFiltering.replace(/script_video_agg_placement_data/g,videoAggregatedPlacementData)
    videoFilteredBadPlacementsSheet.getRange('A1').setFormula(videoFilterFormula)
  }
  
  if(appsModuleEnabled == 'YES'){
    var appsAggregationFormula = '='+appsQueryAggregation.replace(/script_apps_export_from_gads/g,sheetNameReportExportApps)                                                                                 
    appsAggregatedPlacementDataSheet.getRange('A1').setFormula(appsAggregationFormula)   

    var appsFilterFormula = '='+appsQueryFiltering.replace(/script_apps_agg_placement_data/g,appsAggregatedPlacementData)
    appsFilteredBadPlacementsSheet.getRange('A1').setFormula(appsFilterFormula)    


    //Let's create the mobile URLs
    appsFilteredBadPlacementsSheet.getRange('N1').setValue('Normalized Mobile Placement URL')
    appsFilteredBadPlacementsSheet.getRange('N2').setFormula(appsMobileUrlFormula)
  }
  
  //Let's give the spreadsheet sometime to breathe after query formulas are deployed.
  SpreadsheetApp.flush()
  Utilities.sleep(10000)
  Logger.log('Line milestone 520')

  //Let's create the list of placements to negate
  //First I convert array of arrays to a string and then I make an array by splitting it by comma. Finally, I filter out blank cells since via the Boolean trick since B2:B may contain lots of blank cells.
  if(displayModuleEnabled == 'YES'){
    if(displayAnonymousPlacementEnabled == 'YES')
      {var displayBadPLacementsToAdd = displayFilteredBadPlacementsSheet.getRange('B2:B').getValues().join().split(',').filter(Boolean)}
      
    else
      {
       var displayBadPLacementsToAdd = displayFilteredBadPlacementsSheet.getRange('B2:B').getValues().join().split(',').filter(Boolean)
       for(y=0;y<displayBadPLacementsToAdd.length;y++){
         if(displayBadPLacementsToAdd[y]=='anonymous.google'){
           displayBadPLacementsToAdd.splice(y,1)
         }
        }
      }
  }
  
  if(videoModuleEnabled == 'YES'){var videoBadPLacementsToAdd = videoFilteredBadPlacementsSheet.getRange('B2:B').getValues().join().split(',').filter(Boolean)}
  if(appsModuleEnabled == 'YES'){var appsBadPLacementsToAdd = appsFilteredBadPlacementsSheet.getRange('N2:N').getValues().join().split(',').filter(Boolean)}
  
  if(displayModuleEnabled == 'YES'){Logger.log('Bad placements to add - display: '+displayBadPLacementsToAdd)}
  if(videoModuleEnabled == 'YES'){Logger.log('Bad placements to add - video: '+videoBadPLacementsToAdd)}
  if(appsModuleEnabled == 'YES'){Logger.log('Bad placements to add - apps: '+appsBadPLacementsToAdd)}
  
  //Let's do the magic in Google Ads
  //Let's create the list of excluded placements if it does not exist yet.
  
  //Domain endings first
  if(domainEndings.length>0 && domainEndingsModuleEnabled == 'YES'){
    var existingDomainEndingsExcludedPlacements = []
    if(AdsApp.excludedPlacementLists().withCondition("Name = '"+domainEndingsListName+"'").get().totalNumEntities()==1 && domainEndingsModuleEnabled == 'YES')
      {
        Logger.log('List '+domainEndingsListName+' already exists. No need to create.')

        var domainEndingsPlacementsInTheAccount = AdsApp.excludedPlacementLists().withCondition("Name = '"+domainEndingsListName+"'").get().next().excludedPlacements().get()
        while (domainEndingsPlacementsInTheAccount.hasNext()){
          var domainEndingsPlacementInTheAccount = domainEndingsPlacementsInTheAccount.next()

          var domainEndingsPlacementInTheAccountClean = domainEndingsPlacementInTheAccount.toString().replace(/SharedExcludedPlacement: /g,'').replace(/\[/g,'').replace(/\]/g,'')

          existingDomainEndingsExcludedPlacements.push(domainEndingsPlacementInTheAccountClean)
        } 
      }
    else
      {if(domainEndings.length>0 && domainEndingsModuleEnabled == 'YES'){AdsApp.newExcludedPlacementListBuilder().withName(domainEndingsListName).build()}}

    Logger.log('Excluded placements already on the list - domainEndings: '+existingDomainEndingsExcludedPlacements)   
  }
  
 
  //Display
  if(displayModuleEnabled == 'YES'){
    var existingDisplayExcludedPlacements = []

    if(AdsApp.excludedPlacementLists().withCondition("Name = '"+displayListName+"'").get().totalNumEntities()==1 && displayModuleEnabled == 'YES')
      {
        Logger.log('List '+displayListName+' already exists. No need to create.')

        var  displayPlacementsInTheAccount = AdsApp.excludedPlacementLists().withCondition("Name = '"+displayListName+"'").get().next().excludedPlacements().get()
        while (displayPlacementsInTheAccount.hasNext()){
          var displayPlacementInTheAccount = displayPlacementsInTheAccount.next()

          var displayPlacementInTheAccountClean = displayPlacementInTheAccount.toString().replace(/SharedExcludedPlacement: /g,'').replace(/\[/g,'').replace(/\]/g,'')

          existingDisplayExcludedPlacements.push(displayPlacementInTheAccountClean)
        } 
      }
    else
      {if(displayModuleEnabled == 'YES'){AdsApp.newExcludedPlacementListBuilder().withName(displayListName).build()}}

    Logger.log('Excluded placements already on the list - display: '+existingDisplayExcludedPlacements)
  }
  
  //Then video
  if(videoModuleEnabled == 'YES'){
    var existingVideoExcludedPlacements = []

    if(AdsApp.excludedPlacementLists().withCondition("Name = '"+videoListName+"'").get().totalNumEntities()==1  && videoModuleEnabled == 'YES')
      {
        Logger.log('List '+videoListName+' already exists. No need to create.')

        var  videoPlacementsInTheAccount = AdsApp.excludedPlacementLists().withCondition("Name = '"+videoListName+"'").get().next().excludedPlacements().get()
        while (videoPlacementsInTheAccount.hasNext()){
          var videoPlacementInTheAccount = videoPlacementsInTheAccount.next()

          var videoPlacementInTheAccountClean = videoPlacementInTheAccount.toString().replace(/SharedExcludedPlacement: /g,'').replace(/\[/g,'').replace(/\]/g,'')

          existingVideoExcludedPlacements.push(videoPlacementInTheAccountClean)
        } 
      }
    else
      {if(videoModuleEnabled == 'YES'){AdsApp.newExcludedPlacementListBuilder().withName(videoListName).build()}}

    Logger.log('Excluded placements already on the list - video: '+existingVideoExcludedPlacements)    
  }

  //Then apps
  if(appsModuleEnabled == 'YES'){
    
    var existingAppExcludedPlacements = []

    if(AdsApp.excludedPlacementLists().withCondition("Name = '"+appListName+"'").get().totalNumEntities()==1  && appsModuleEnabled == 'YES')
      {
        Logger.log('List '+appListName+' already exists. No need to create.')

        var  appPlacementsInTheAccount = AdsApp.excludedPlacementLists().withCondition("Name = '"+appListName+"'").get().next().excludedPlacements().get()
        while (appPlacementsInTheAccount.hasNext()){
          var appPlacementInTheAccount = appPlacementsInTheAccount.next()

          var appPlacementInTheAccountClean = appPlacementInTheAccount.toString().replace(/SharedExcludedPlacement: /g,'').replace(/\[/g,'').replace(/\]/g,'')

          existingAppExcludedPlacements.push(appPlacementInTheAccountClean)
        } 
      }
    else
      {if(appsModuleEnabled == 'YES'){AdsApp.newExcludedPlacementListBuilder().withName(appListName).build()}}

    Logger.log('Excluded placements already on the list - app: '+existingAppExcludedPlacements)   
  }

  Logger.log('Line milestone 642')

  //Let's find the placements which are not on our list yet
  //Domain Endings first
  
  if(domainEndings.length>0 && domainEndingsModuleEnabled == 'YES'){
    Logger.log('')
    Logger.log('***Starting comparison of existing vs new placements - domain endings***')
    Logger.log('existingDomainEndingsExcludedPlacements: '+existingDomainEndingsExcludedPlacements)
    Logger.log('domainEndingsBadPLacementsToAdd: '+placementsToBeExcludedBasedOnEndings)

    var finalListOfDomainEndingsPlacementsToAdd = []

    for(i=0;i<placementsToBeExcludedBasedOnEndingsUnique.length;i++){
      if(existingDomainEndingsExcludedPlacements.indexOf(placementsToBeExcludedBasedOnEndingsUnique[i])==-1) //-1 means not found
         {  
         finalListOfDomainEndingsPlacementsToAdd.push(placementsToBeExcludedBasedOnEndingsUnique[i]) 
         if(loggerSetup=='HEAVY'){Logger.log(placementsToBeExcludedBasedOnEndingsUnique[i]+' index: '+existingDomainEndingsExcludedPlacements.indexOf(placementsToBeExcludedBasedOnEndingsUnique[i]))}
         }
      else
        {
        if(loggerSetup=='HEAVY'){Logger.log(placementsToBeExcludedBasedOnEndingsUnique[i]+' is already on the list. No need to add again. Index: '+existingDomainEndingsExcludedPlacements.indexOf(placementsToBeExcludedBasedOnEndingsUnique[i]))}
        }
    }

    Logger.log('Final list of bad placements to add: '+finalListOfDomainEndingsPlacementsToAdd)    
    Logger.log('**********End of comparison - domain endings*************')
    Logger.log('')    
  }
  
  Logger.log('Line milestone 672')

  //Display
  if(displayModuleEnabled == 'YES'){
    Logger.log('')
    Logger.log('***Starting comparison of existing vs new placements - display***')
    Logger.log('existingDisplayExcludedPlacements: '+existingDisplayExcludedPlacements)
    Logger.log('displayBadPLacementsToAdd: '+displayBadPLacementsToAdd)

    var finalListOfDisplayPlacementsToAdd = []

    for(i=0;i<displayBadPLacementsToAdd.length;i++){
      if(existingDisplayExcludedPlacements.indexOf(displayBadPLacementsToAdd[i])==-1) //-1 means not found
         {  
         finalListOfDisplayPlacementsToAdd.push(displayBadPLacementsToAdd[i]) 
         if(loggerSetup=='HEAVY'){Logger.log(displayBadPLacementsToAdd[i]+' index: '+existingDisplayExcludedPlacements.indexOf(displayBadPLacementsToAdd[i]))}
         }
      else
        {
        if(loggerSetup=='HEAVY'){Logger.log(displayBadPLacementsToAdd[i]+' is already on the list. No need to add again. Index: '+existingDisplayExcludedPlacements.indexOf(displayBadPLacementsToAdd[i]))}
        }
    }

    Logger.log('Final list of bad placements to add: '+finalListOfDisplayPlacementsToAdd)    
    Logger.log('**********End of comparison - display*************')
    Logger.log('')  
  }

  Logger.log('Line milestone 700')

  //Then video
  if(videoModuleEnabled == 'YES'){
    Logger.log('')
    Logger.log('***Starting comparison of existing vs new placements - video***')
    Logger.log('existingVideoExcludedPlacements: '+existingVideoExcludedPlacements)
    Logger.log('videoBadPLacementsToAdd: '+videoBadPLacementsToAdd)

    var finalListOfVideoPlacementsToAdd = []

    for(i=0;i<videoBadPLacementsToAdd.length;i++){
      if(existingVideoExcludedPlacements.indexOf(videoBadPLacementsToAdd[i])==-1) //-1 means not found
         {  
         finalListOfVideoPlacementsToAdd.push(videoBadPLacementsToAdd[i]) 
         if(loggerSetup == 'HEAVY'){Logger.log(videoBadPLacementsToAdd[i]+' index: '+existingVideoExcludedPlacements.indexOf(videoBadPLacementsToAdd[i]))}
         }
      else
        {
        if(loggerSetup=='HEAVY'){Logger.log(videoBadPLacementsToAdd[i]+' is already on the list. No need to add again. Index: '+existingVideoExcludedPlacements.indexOf(videoBadPLacementsToAdd[i]))}
        }
    }

    Logger.log('Final list of bad placements to add: '+finalListOfVideoPlacementsToAdd)    
    Logger.log('**********End of comparison - video*************')
    Logger.log('')  
  }
  
  Logger.log('Line milestone 728')

  //Then app
  if(appsModuleEnabled == 'YES'){
    Logger.log('')
    Logger.log('***Starting comparison of existing vs new placements - app***')
    Logger.log('existingAppExcludedPlacements: '+existingAppExcludedPlacements)
    Logger.log('appBadPLacementsToAdd: '+appsBadPLacementsToAdd)
  
    var finalListOfAppPlacementsToAdd = []

    for(i=0;i<appsBadPLacementsToAdd.length;i++){
      if(existingAppExcludedPlacements.indexOf(appsBadPLacementsToAdd[i])==-1) //-1 means not found
         {  
         finalListOfAppPlacementsToAdd.push(appsBadPLacementsToAdd[i]) 
         if(loggerSetup=='HEAVY'){Logger.log(appsBadPLacementsToAdd[i]+' index: '+existingAppExcludedPlacements.indexOf(appsBadPLacementsToAdd[i]))}
         }
      else
        {
        if(loggerSetup=='HEAVY'){Logger.log(appsBadPLacementsToAdd[i]+' is already on the list. No need to add again. Index: '+existingAppExcludedPlacements.indexOf(appsBadPLacementsToAdd[i]))}
        }
    }

    Logger.log('Final list of bad placements to add: '+finalListOfAppPlacementsToAdd)    
    Logger.log('**********End of comparison - app*************')
    Logger.log('')    
  
  }
  
  Logger.log('Line milestone 757')

  //Let's add the latest shit placements into the list - not big deal if you add the same placements over and over since the script won't fail.
  //But only if the script is set to run for real
  if(forRealOrTesting=='RUNNING_FOR_REAL' && domainEndingsModuleEnabled == 'YES' && domainEndings.length>0){AdsApp.excludedPlacementLists().withCondition("Name = '"+domainEndingsListName+"'").get().next().addExcludedPlacements(finalListOfDomainEndingsPlacementsToAdd)} 
  if(forRealOrTesting=='RUNNING_FOR_REAL' && displayModuleEnabled == 'YES'){AdsApp.excludedPlacementLists().withCondition("Name = '"+displayListName+"'").get().next().addExcludedPlacements(finalListOfDisplayPlacementsToAdd)} 
  if(forRealOrTesting=='RUNNING_FOR_REAL' && videoModuleEnabled == 'YES'){AdsApp.excludedPlacementLists().withCondition("Name = '"+videoListName+"'").get().next().addExcludedPlacements(finalListOfVideoPlacementsToAdd)} 
  if(forRealOrTesting=='RUNNING_FOR_REAL' && appsModuleEnabled == 'YES'){AdsApp.excludedPlacementLists().withCondition("Name = '"+appListName+"'").get().next().addExcludedPlacements(finalListOfAppPlacementsToAdd)} 
  
  //Let's assign Display campaigns with negative lists
  var displayCampaignsWhichGotDisplayExclusionList = []
  var displayCampaignsWhichGotAppsExclusionList = []
  var displayCampaignsWhichGotDomainEndingsExclusionList = []
  
  Logger.log('Line milestone 771')

  //The check for module enabled is down in the code - it's desirable here
  for(i=0;i<displayCampaignNames.length;i++){  

    var displayCampaigns = AdsApp.campaigns().withCondition("Status = 'ENABLED' AND Name = '"+displayCampaignNames[i]+"'").get()

    while(displayCampaigns.hasNext()){
      var displayCampaign = displayCampaigns.next()

      //Display plus app list
      //The total Num Entities trick is there to check whether the list is already assigned or not - if not then I am assigning the list to the campaign
      if(displayModuleEnabled == 'YES' && displayCampaign.excludedPlacementLists().withCondition("Name = '"+displayListName+"'").get().totalNumEntities()!=1){
        displayCampaign.addExcludedPlacementList(AdsApp.excludedPlacementLists().withCondition("Name = '"+displayListName+"'").get().next())
        displayCampaignsWhichGotDisplayExclusionList.push(displayCampaign)
      }
      if(appsModuleEnabled == 'YES' && displayCampaign.excludedPlacementLists().withCondition("Name = '"+appListName+"'").get().totalNumEntities()!=1){
        displayCampaign.addExcludedPlacementList(AdsApp.excludedPlacementLists().withCondition("Name = '"+appListName+"'").get().next())
        displayCampaignsWhichGotAppsExclusionList.push(displayCampaign)
      }
      if(domainEndingsModuleEnabled == 'YES' && domainEndings.length>0 && displayCampaign.excludedPlacementLists().withCondition("Name = '"+domainEndingsListName+"'").get().totalNumEntities()!=1){
        displayCampaign.addExcludedPlacementList(AdsApp.excludedPlacementLists().withCondition("Name = '"+domainEndingsListName+"'").get().next())
        displayCampaignsWhichGotDomainEndingsExclusionList.push(displayCampaign)
      }          
    }
  }
  
  Logger.log('Line milestone 798')
 
  //Now video campaigns - assign them the video list]
  var videoCampaignsWhichGotVideoExclusionList = []
  var videoCampaignsWhichGotDomainEndingsExclusionList = []
  
  
  for(s=0;s<videoCampaignNames.length;s++){ 

    //Let's take care of bumper campaigns
    if(videoBumperExclusion == 'YES' && videoBumperString.length>0)
    {var videoCampaigns = AdsApp.campaigns().withCondition("Status = 'ENABLED' AND Name = '"+videoCampaignNames[i]+"' AND Name DOES_NOT_CONTAIN '"+videoBumperString+"'").get()}    
    else
    {var videoCampaigns = AdsApp.campaigns().withCondition("Status = 'ENABLED' AND Name = '"+videoCampaignNames[i]+"'").get()}

    while(videoCampaigns.hasNext()){
      var videoCampaign = videoCampaigns.next()      

      //If the list is not assigned yet, then assign...
      if(videoModuleEnabled == 'YES' && videoCampaign.excludedPlacementLists().withCondition("Name = '"+videoListName+"'").get().totalNumEntities()!=1){
        videoCampaign.addExcludedPlacementList(AdsApp.excludedPlacementLists().withCondition("Name = '"+videoListName+"'").get().next())
        videoCampaignsWhichGotVideoExclusionList.push(videoCampaign)
      }

      if(domainEndingsModuleEnabled == 'YES' && domainEndings.length>0 && videoCampaign.excludedPlacementLists().withCondition("Name = '"+domainEndingsListName+"'").get().totalNumEntities()!=1){
        videoCampaign.addExcludedPlacementList(AdsApp.excludedPlacementLists().withCondition("Name = '"+domainEndingsListName+"'").get().next())
        videoCampaignsWhichGotDomainEndingsExclusionList.push(videoCampaign)
      }    
    }
  }
    
  

  /* The idea was to also assign exlusions lists to smart shopping campaigns but they are not suported in Google scripts yet, so commenting out for now
  //Now shopping campaigns - give them the display and video list
  var shoppingCampaigns = AdsApp.shoppingCampaigns().withCondition("Status = 'ENABLED'").get()
  
  while(shoppingCampaigns.hasNext()){
    var eachEnabledShoppingCampaign = shoppingCampaigns.next()
    
    //There's no check whether the list has been already assigned or not. Google Ads will not throw an error if you try to assign the list again.
    //Display+video+app negative list
    
    if(displayModuleEnabled == 'YES'){eachEnabledShoppingCampaign.addExcludedPlacementList(AdsApp.excludedPlacementLists().withCondition("Name = '"+displayListName+"'").get().next())}
    if(videoModuleEnabled == 'YES'){eachEnabledShoppingCampaign.addExcludedPlacementList(AdsApp.excludedPlacementLists().withCondition("Name = '"+videoListName+"'").get().next())}
    if(appsModuleEnabled == 'YES'){eachEnabledShoppingCampaign.addExcludedPlacementList(AdsApp.excludedPlacementLists().withCondition("Name = '"+appListName+"'").get().next())}

  }
  */
  
  Logger.log('Line milestone 848')

  //Let's send the email
  var emailPart = UrlFetchApp.fetch('https://drive.google.com/uc?export=view&id=15f90xa2Coga_a7kctn7Wk6jdKRc2LT5W').getContentText()  
  eval(emailPart)
  
  //TEMP EMAIL****
  //Paste here....
  //END OF TEMP***
  
  
  Logger.log('Script finished.')
  
  //send event completed to GA
  var payload = {
    'v': '1',
    'tid': 'UA-69459605-1',
    't': 'event',
    'cid' : Utilities.getUuid(),
    'z' : Math.floor(Math.random() * 10E7),
    'ds' : 'google_ads_script',
    'cn' : 'script-bad_placement_cleaner',
    'cs' : 'google_ads_script',
    'cm' : 'google_ads_script',
    'ec' : 'script_run',
    'ea' : 'script_run_completed',
    'el' : 'script-bad_placement_cleaner'
    };
  
  var options = {
    "method" : "post",
    "payload" : payload
   };
  
  // Send the hit to GA if allowed
  if(eventReportingToExcelinPpcComAllowed == 'YES'){UrlFetchApp.fetch("http://www.google-analytics.com/collect", options)};
  
  Logger.log('You can join my FB group so you don\'t miss any cool scripts in future! Join here: https://www.facebook.com/groups/Excelinppc/')  
  Logger.log('If something is not working, visit http://tinyurl.com/tkqkotq and submit a comment under the article.')