'******************************************************************************************************************************
' Name Of Business Component		:	Start Active Workspace
'
' Purpose							:	Invoke the application and Login to Active Workspace
'
' Input	Parameter					:	Parameter 1: str_Instance 
'
'										Parameter 2: str_UserName 
'
'										Parameter 3: str_Password 
'
' Output							:	True / False
'
' Remarks							:
'
' Author							:	Mohini Deshmukh 			  29 July 2020

'******************************************************************************************************************************
Option Explicit

Dim testPath, resourcesParentFolder
Dim resourceName, repoFullPath
resourceName = "AWC_OR_ChromiumEdge.tsr"
testPath =Environment("TestDir")

resourcesParentFolder = FindParentFolderWithResources(testPath)
If resourcesParentFolder = "" Then
	Reporter.ReportEvent micFail, "Error searching Test Resource: " & filename, "Couldn't find the folder containing TestResources."
End If
resourcesParentFolder = resourcesParentFolder & "\TestResources"
repoFullPath = FindResourceFullPath(resourcesParentFolder, resourceName)

Dim repos
Set repos = RepositoriesCollection
repos.removeAll

repos.Add repoFullPath


'-------------------------------------------------------------------------------------------------------------------------------
'Variable Declaration
'-------------------------------------------------------------------------------------------------------------------------------
Dim objEdgeBrowser,objBrowsers,objApp,obj_AWCTeamcenterHome,Processes,objBrowsersNew,objWshShell
Dim iCount,iCounter
Dim sVersion,sGroup,sRole,sTempValue,Process,targetUrl,sPassword,sUserName,sHeaderText,sTemp,sTempUserNm
Dim myProcess1,myProcess


Reporter.ReportEvent micPass, "Test dir", Environment("TestDir")

Dim repoCount, i
repoCount = RepositoriesCollection.Count

For i = 1 To repoCount
    Reporter.ReportEvent micPass, "FOUND OBJECT REPOSITORY:", "Repository " & i & ": " & RepositoriesCollection.Item(i)
Next

'--------------------------------------------------------------------------------------------------------------------------------
'Get AWC PLM window object from xml
'--------------------------------------------------------------------------------------------------------------------------------
Set obj_AWCTeamcenterHome=Eval(SearchAndLoadResourceByName("ActiveWorkspace_OR.xml").GetValue("wpage_AWCTeamcenterHome"))
'Set obj_AWCTeamcenterHome=Browser("Browser").Page("wpage_AWCTeamcenterHome")


Set obj_AWCTeamcenterHome = Nothing
