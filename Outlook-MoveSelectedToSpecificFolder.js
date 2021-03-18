#!/usr/bin/env osascript -l JavaScript
const outlook = Application('Microsoft Outlook');
const DEBUG = true;
// although there is ES6 support , but "run" function
// can't use fat arrow functions.
// https://github.com/JXA-Cookbook/JXA-Cookbook/wiki/ES6-Features-in-JXA#arrow-functions
function run(args) {
  
  const DST_FOLDER = getKMVar("local_destinationFolder");
  
  
  
  var selectedMessageList = outlook.currentMessages();
  
  
  var selectedMessageSubject = outlook.selectedObjects()[0].subject();
  var app = Application.currentApplication();
  app.includeStandardAdditions = true;
  
  //console.log(outlook.selectedObjects()[0].subject())

  for (let idxMessage = 0 ; idxMessage < selectedMessageList.length; idxMessage++)
  {
    
    // Get the Account in which the Message we are processing resides
    let currentAccount = selectedMessageList[idxMessage].account();
    DEBUG && console.log("Account for next move:", currentAccount.name());
    
    // Search for our Destination Folder in current Account
    let dstFolder = currentAccount.mailFolders().find(function (elem) {
      return elem.name() == DST_FOLDER;
    });


    DEBUG && console.log("Moving eMail: ", selectedMessageList[idxMessage].subject());
    
    // Let's move the message
    outlook.move(selectedMessageList[idxMessage], { to: dstFolder });
    
    // display a Notification
    selectedMessageSubject = selectedMessageList[idxMessage].subject();
    app.displayNotification("Moved: " + selectedMessageSubject);

  }
  
  
}



//=====================================================================	
function getKMVar(pstrName) {
//=====================================================================	
  
  
  
  var app = Application.currentApplication()
  app.includeStandardAdditions = true
  
  var kmInst = app.systemAttribute("KMINSTANCE");
  var kmeApp = Application("Keyboard Maestro Engine")
  

  var myLocalVar = kmeApp.getvariable(pstrName,  {instance: kmInst});
  //kmeApp.setvariable("Local__FromJXA", {instance: kmInst, to: "Set in JXA Script"})
  
  
  return myLocalVar
  
}	// END function getKMVar