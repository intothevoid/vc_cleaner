' VC++ Cleaner - VB Script to clean intermediate files from a VC++ project
' Copyright (C) 2010-2014 Karan Kadam
' License: http://www.gnu.org/licenses/gpl.html GPL version 2 or higher

'Declarations
OPTION EXPLICIT
DIM strExt,strDir, strDirDel, strCurrExtn, strLogFile
DIM objFSO, SubDir, objTextFile

'Resume and dont display errors
'ON ERROR RESUME NEXT

'List of extensions to be deleted
'Comma seperated values. No spaces
strExt = "pch,clw,aps,plg,opt,ncb,scc,ilk,htm,vss,pcx,bkp,bak,bsc,user,suo"

'List of folders to be deleted
'example - delete the debug, release folders
strDirDel = "debug,release"

'Display banner
wscript.echo "VC++ Cleaner BETA"

'Set logfile name
strLogFile = "vcpp_cleaner.log"

'Create filesystem object
SET objFSO = CREATEOBJECT("Scripting.FileSystemObject")

'Create log file
SET objTextFile = objFSO.OpenTextFile(strLogFile, 2, True) '2 = for writing

'Path to the folder from which files are to be deleted
strDir = Inputbox("Enter path to clean - ","VC++ Cleaner")

'Delete files from sub-directories
SubDir = TRUE

'Log start
objTextFile.WriteLine("VC++ Cleaner - Started")
	
'Delete files
DeleteFiles strDir,strExt,strDirDel,SubDir,objTextFile

'Delete folders
DeleteFolders strDir,strDirDel,objTextFile

'Log complete
objTextFile.WriteLine("VC++ Cleaner - Finished")

'Close log file
objTextFile.Close

'Deletion complete
wscript.echo "VC++ Cleaner - Finished"

'This function is used for cleaning all the intermediate files
SUB DeleteFiles(BYVAL strDirectory,BYVAL strExt,BYVAL strDirDel,SubDir,objTextFile)
	DIM objFolder, objSubFolder, objFile
	DIM strBuff
		
	'Set directory
	SET objFolder = objFSO.GetFolder(strDirectory)
		
	FOR EACH objFile in objFolder.Files
	    'wscript.echo "Debug (Curr Filename):" & objFile.Path
		FOR EACH strCurrExtn in SPLIT(UCASE(strExt),",")
			IF RIGHT(UCASE(objFile.Path),LEN(strCurrExtn)+1) = "." & strCurrExtn THEN
					strBuff = "Deleting file: " & objFile.Path & " | " & objFile.DateLastModified 
					objTextFile.WriteLine(strBuff)
					objFile.Delete
					EXIT FOR
			END IF
		NEXT
	NEXT	
	
	'If subdirectories present, then do a recursive delete
	IF SubDir = TRUE THEN 
		FOR EACH objSubFolder in objFolder.SubFolders
			DeleteFiles objSubFolder.Path,strExt,strDirDel,SubDir,objTextFile
		NEXT
	END IF
END SUB

'This function is used for cleaning all the intermediate folders
SUB DeleteFolders(BYVAL strDirectory,BYVAL strDirDel,objTextFile)
    DIM objFolder, objSubFolder
    DIM strTmp,strTmp2, strCurrDir
    DIM intCount,nCnt
    DIM arrFolders
  
    'Init array
    arrFolders = Array()
    
    'Set directory
	SET objFolder = objFSO.GetFolder(strDirectory)

    'Check if intermediate folders found in current level    
    FOR EACH objSubFolder in objFolder.SubFolders
        FOR EACH strCurrDir in SPLIT(UCASE(strDirDel),",")
            strTmp = (RIGHT(UCASE(objSubFolder.Path),LEN(strCurrDir)))
            'objTextFile.WriteLine(strTmp)
            'objTextFile.WriteLine(strCurrDir)
            
            IF strTmp = strCurrDir THEN
                strTmp2 = "Deleting folder: " & objSubFolder.Path & " | " & objSubFolder.DateLastModified 
                objTextFile.WriteLine(strTmp2)
                intCount = UBound(arrFolders) + 1
                ReDim Preserve arrFolders(intCount)
                arrFolders(intCount) = objSubFolder.Path
            END IF            
        NEXT
    NEXT
   
    'Delete all folders found on current level    
    For nCnt = 0 To UBound(arrFolders)
        objFSO.DeleteFolder arrFolders(nCnt), True
    Next
    
    'Check in sub directories
    FOR EACH objSubFolder in objFolder.SubFolders
        'Recursive call for other folders
        DeleteFolders objSubFolder.Path,strDirDel,objTextFile    
    NEXT
END SUB
