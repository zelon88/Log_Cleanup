'File Name: Logon.vbs
'Version: v1.0, 1/26/2013
'Author: Justin Grimes, 1/26/2023

' --------------------------------------------------
'Configure the execution environment for the script.
Option Explicit
'On Error Resume Next
Dim fileSystem, objFso, oShell, userName, computerName, strSafeDate, strSafeDateA, strSafeTime, strSafeTimeA, strDateTime, logfile, objFolder, companyName, result, _ 
  objFile, strFolderPath, intDaysOlderThan, outputData, error, startMessage, message, dataMessage, intMessage, fileList, logPath, appPath, companyAbbr, objLogfile
' --------------------------------------------------

' --------------------------------------------------
'Declare Objects.
Set objFso = CreateObject("Scripting.FileSystemObject")
Set oShell = WScript.CreateObject("WScript.Shell")
Set fileSystem = CreateObject("Scripting.FileSystemObject")
' --------------------------------------------------

' --------------------------------------------------
'Company Specific variables.
'Change the following variables to match the details of your organization.

'The "appPath" is the full absolute path for the script directory, with trailing slash.
appPath = "\\SERVER\Scripts\Logon\"
'The "logPath" is the full absolute path for where network-wide logs are stored.
logPath = "\\SERVER\Logs"
'The "companyName" the the full, unabbreviated name of your organization.
companyName = "Company Inc."
'The "companyAbbr" is the abbreviated name of your organization.
companyAbbr = "Company"
'The "strFolderPath" is the location to look for stale files.
strFolderPath = "\\SERVER\Logs"
'The "intDaysOlderThan" is the number of days older than which a file will be deleted.
intDaysOlderThan = 28
' --------------------------------------------------

' --------------------------------------------------
'Declare Variables.
userName = CreateObject("WScript.Network").UserName
computerName = oShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
strSafeDate = DatePart("yyyy",Date)&Right("0"&DatePart("m",Date), 2)&Right("0"&DatePart("d",Date), 2)
strSafeDateA = DatePart("yyyy",Date)&"-"&Right("0"&DatePart("m",Date), 2)&"-"&Right("0"&DatePart("d",Date), 2)
strSafeTime = Right("0"&Hour(Now), 2)&Right("0"&Minute(Now), 2)&Right("0"&Second(Now), 2)
strSafeTimeA = Right("0"&Hour(Now), 2)&":"&Right("0"&Minute(Now), 2)&":"&Right("0"&Second(Now), 2)
strDateTime = strSafeDate&"-"&strSafeTime
logfile = logPath&"\"&computerName&"-"&userName&"-"&strDateTime&"-Log_Cleanup.txt"
startMessage = "The user "&userName&" has cleaned the log directory from the system "&computerName&" on "&strSafeDateA&" at "&strSafeTimeA&"."&vbNewLine
dataMessage = "The following files were removed:"&vbNewLine
intMessage = startMessage&dataMessage
' --------------------------------------------------

' --------------------------------------------------
'A function to create a log file is set.
'Returns "True" if logfile exists, "False" on error.
Function CreateLog(logfile, message)
  If message <> "" Then
    Set objLogfile = fileSystem.CreateTextFile(logfile, True)
    objLogfile.WriteLine(message)
    objLogfile.Close
  End If
End Function
' --------------------------------------------------

' --------------------------------------------------
'A function to delete files contained within a given path that are older than a specified number of days.
Function DeleteFiles(path, days)
  Set objFolder = objFso.GetFolder(path)
  outputData = vbNewLine
  For Each objFile In objFolder.Files
    If objFile.DateLastModified < (Now() - days) Then
      outputData = outputData&objFso.GetAbsolutePathName(objFile)&vbNewLine
      objFile.Delete(True)
    End If
  Next
  DeleteFiles = outputData
End Function
' --------------------------------------------------

' --------------------------------------------------
'The main logic of the script which makes use of the functions above.
fileList = DeleteFiles(strFolderPath, intDaysOlderThan)
message = intMessage&fileList
CreateLog logfile, message
' --------------------------------------------------
