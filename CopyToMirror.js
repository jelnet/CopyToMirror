/*
Author: Jeremy Wray, FM&T Vision 
Version 2.0 
*/

//create objects
var objArgs = WScript.Arguments;
var objFSO = WScript.CreateObject("Scripting.FileSystemObject");
var objS = WScript.CreateObject("WScript.Shell");
//app file name 
var strAppFileName = "CopyToMirror.js";
//app name and version
// 1.1 updated getMirror func to allow network paths not just drive mappings
// 2.0 updated to new vdev/redesign/stage environment
var strAppName = "CopyToMirror 2.0";
//boolean for overwriting existing files/folders
var blnOverwrite = false;


//check someone's right-clicked a file, if not see if they want to install the app
if (objArgs.length == 0){

        var blnInstall = objS.Popup("Install " + strAppName + "?",-1,strAppName,36);      
        
        //if yes
        if (blnInstall == 6){
      
            //get current location of app file
            var source_obj = objS.CurrentDirectory + "\\" + strAppFileName;       
            //get location of user's Send To menu 
            var target_obj = objS.SpecialFolders("SendTo") + "\\";
            //boolean for overwriting existing app
             var blnOverwrite = false;
           
            //function to try copying the file 
            tryCopyApp = function() {                
                try {
                     objFSO.CopyFile (source_obj, target_obj, blnOverwrite);
                }
                 //if already exists get user confirmation else quit
                catch (e) {    
                    if (e.message == "File already exists"){           
                    blnOverwrite = objS.Popup("App already exists and will be overwritten\nProceed?",-1,strAppName,49);
                        if (blnOverwrite == 1){
                            tryCopyApp(); return;         
                            }else{
                            WScript.Quit(); 
                        }                        
                    }else{
                     //show any other errors and show retry button else quit
                    if (objS.Popup(e.message,-1,strAppName,21) == 4){   
                            tryCopyApp(); return;       
                        }else{
                            WScript.Quit();
                        }
                    }       
                }
                //if we get here install was successful
                objS.Popup(strAppName + " installed to " + target_obj + ".\nDeploy by right-clicking a folder or file you want to copy from dev\\redesign to stage and selecting '" + strAppFileName + "' from the Send To menu.",-1,strAppName,64);
            }
                            
            tryCopyApp();
            
         }         

     WScript.Quit();
};


//Get parent folder of intended target, only need this once 
var target_path = objFSO.GetParentFolderName(getMirror(objArgs(0)));

//Check that target is valid (ie. are within dev or test subfolders */
if (! target_path){
    objS.Popup("Item does not appear to be within a dev, redesign or stage subdirectory",-1,strAppName,16);
    WScript.Quit();
}




//loop through passed args (files/folders)
for(var i=0; i<objArgs.length; i++){
    
    //get full path and name of passed file/folder
    var source_obj = objArgs(i);    
   //get full path and name of intended target (eg. map dev to test and vice-versa) of passed file/folder   
    var target_obj = getMirror(source_obj);      
    
     //if argument is a folder
     if (objFSO.FolderExists(source_obj)){
        
         //see if the path to the target file exists, if not build a path to it                         
        buildPath(target_path);                 
        
        //function to try copying the folder 
        tryCopyFolder = function() {
            try {
                objFSO.CopyFolder (source_obj, target_obj, blnOverwrite);
            }
            //if already exists get user confirmation else quit
            catch (e) {    
                if (e.message == "File already exists"){           
                blnOverwrite = objS.Popup("Folder(s) will be overwritten on:\n" + target_path + "\nProceed?",-1,strAppName,49);
                    if (blnOverwrite == 1){
                        tryCopyFolder(); return;              
                        }else{
                        WScript.Quit(); 
                    }                        
                }else{
                //show any other errors and show retry button else quit
                if (objS.Popup(e.message,-1,strAppName,21) == 4){   
                        tryCopyFolder(); return;            
                    }else{
                        WScript.Quit();
                    }
                }       
            }
        }
        
        tryCopyFolder();
     
        
    //if argument is a file
    }else if (objFSO.FileExists(source_obj)){
    
        //see if the path to the target file exists, if not build a path to it                         
        buildPath(target_path);      
        
        //function to try copying the file 
        tryCopyFile = function() {
            try {
                objFSO.CopyFile (source_obj, target_obj, blnOverwrite);
            }
             //if already exists get user confirmation else quit
            catch (e) {    
                if (e.message == "File already exists"){           
                blnOverwrite = objS.Popup("File(s) will be overwritten on:\n" + target_path + "\nProceed?",-1,strAppName,49);
                    if (blnOverwrite == 1){
                        tryCopyFile(); return;                    
                        }else{
                        WScript.Quit(); 
                    }                        
                }else{
                 //show any other errors and show retry button else quit
                if (objS.Popup(e.message,-1,strAppName,21) == 4){   
                        tryCopyFile(); return;           
                    }else{
                        WScript.Quit();
                    }
                }       
            }
        }
                        
        tryCopyFile();
        
       
     //if argument is not folder or file      
    }else{
        WScript.Echo("Object not found");
        WScript.Quit();
    }
 }
 
 //function to create path to target file/dir
 function buildPath(path){
 
    //see if the path to the target folder exists
    if (objFSO.FolderExists(path)){     
        return true;            
    }    
    
    //if not build a path to it:      
    var path_temp = path;
    var absent_folders = [];
    var count = 0;
    
        /*go back though file path till we find a folder that exists, 
        all the while storing missing folder names in array */
        while (! objFSO.FolderExists(path_temp)){        
            absent_folders[count]=objFSO.GetBaseName(path_temp);
            path_temp = objFSO.GetParentFolderName(path_temp);  
             //to prevent hanging in case of non-existent mirror
            if (count > 100){break;}                        
            count++;                 
        }
         
        /*go back through missing folder names array and create folders
        starting from the point where folder does exists (found above) */
        for (var j=0; j<absent_folders.length; j++){       
            count--;            
            try {
                objFSO.CreateFolder(path_temp += "\\" + absent_folders[count]);      
            }
            catch (e) {
                objS.Popup(e.message,-1,strAppName,16);
                WScript.Quit();    
            }     
        }  
    
    //return now-existing path if successfully created    
    if ( objFSO.FolderExists(path)){
        return objFSO.FolderExists(path);
        }else{
        WScript.Echo("There was a problem creating path to: " + path + "\nQuitting");
        WScript.Quit();     
        return false; 
    }  
 }

//show message to confirm number items copied and to allow user to explore items
var blnExp = objS.Popup(objArgs.length + " item(s) copied to:\n" + target_path + "\nExplore?",-1,strAppName,36);

//if user wants to explore open Explorer with target path
if (blnExp == 6){   
    objS.Exec("explorer " +  target_path);
}

/*function to return file mirror path from dev to test server and vice-versa. 
or false if not valid */
function getMirror(loc){    
    if (loc.match(/^.+\\stage\\/)){        
        return loc.replace(/stage/,"dev");
    }else if (loc.match(/^.+\\dev\\/)){            
        return loc.replace(/dev/,"stage");
	}else if (loc.match(/^.+\\_test\\/)){        
        return loc.replace(/_test/,"_dev");
    }else if (loc.match(/^.+\\_dev\\/)){            
        return loc.replace(/_dev/,"_test");
	}else if (loc.match(/^.+\\redesign\\/)){            
        return loc.replace(/redesign/,"stage");
    }else{       
         return false;
     }
}