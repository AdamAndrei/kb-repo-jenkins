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
'-------------------------------------------------------------------------------------------------------------------------------
'Variable Declaration
'-------------------------------------------------------------------------------------------------------------------------------
Dim objEdgeBrowser,objBrowsers,objApp,obj_AWCTeamcenterHome,Processes,objBrowsersNew,objWshShell
Dim iCount,iCounter
Dim sVersion,sGroup,sRole,sTempValue,Process,targetUrl,sPassword,sUserName,sHeaderText,sTemp,sTempUserNm
Dim myProcess1,myProcess

Dim repos
'Set repos = ObjectRepositories
' Add a shared object repository by its full path
'RepositoriesCollection.Add "..\..\..\Resources\BPT Resources\Object Repositories\AWC_OR_ChromiumEdge.tsr"


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
Set objWshShell = CreateObject("WScript.shell")
'--------------------------------------------------------------------------------------------------------------------------------
'Terminate all sprinter processes 
'--------------------------------------------------------------------------------------------------------------------------------
myProcess="Sprinter.exe"    
myProcess1="SprinterAgent.exe" 

Set Processes = GetObject("winmgmts:").InstancesOf("Win32_Process")

For Each Process In Processes
	'--------------------------------------------------------------------------------------------------------------------------------
	'Check the sprinter process if exist then terminate it 
	'--------------------------------------------------------------------------------------------------------------------------------
	  If StrComp(Process.Name, myProcess, vbTextCompare) = 0 Then 
	      Process.Terminate()                                   
	  End If
	    '--------------------------------------------------------------------------------------------------------------------------------
	    'Check the sprinteragent process if exist then terminate it
	    '--------------------------------------------------------------------------------------------------------------------------------
	    If StrComp(Process.Name, myProcess1, vbTextCompare) = 0 Then 
	       Process.Terminate()                                   
	    End If
	 
Next
'--------------------------------------------------------------------------------------------------------------------------------
'Creating Description object
'--------------------------------------------------------------------------------------------------------------------------------
Set objEdgeBrowser = Description.Create()
objEdgeBrowser("micclass").Value = "Browser"
objEdgeBrowser("version").Value="Chromium Edge.*"

sVersion="version:=Chromium Edge.*"

'Set objBrowsers = desktop.ChildObjects(objEdgeBrowser)
'wait 1
'For iCount =0 to objBrowsers.count-1

	Set objBrowsersNew = desktop.ChildObjects(objEdgeBrowser)	
	If objBrowsersNew.count>0 Then
		If obj_AWCTeamcenterHome.exist Then
			If obj_AWCTeamcenterHome.WebButton("wbtn_Your_Profile").exist Then
	 			obj_AWCTeamcenterHome.WebButton("wbtn_Your_Profile").Click
	 			wait 2
		 		If Fn_Web_UI_WebElement_Operations("Signout from Active Workspace application","Click",obj_AWCTeamcenterHome.WebButton("wbtn_SignOut"),"","","","")=True Then
					Reporter.ReportEvent micPass, "Click on [ Sign out ] button in Active Workspace", "Successfully Clicked on [ Sign out ] button in Active Workspace"
				Else
					Reporter.ReportEvent micFail, "Click on [ Sign out ] button in Active Workspace", "Fail to click on [ Sign out ] button in Active Workspace"
					ExitComponent
				End  If
			End  IF	
			SystemUtil.CloseProcessByName "msedge.exe"
		Else
			SystemUtil.CloseProcessByName "msedge.exe"
		End  IF	
		wait 2
	End  IF	

Set Processes = GetObject("winmgmts:").InstancesOf("Win32_Process")
myProcess="msedge.exe"
For Each Process In Processes
	'--------------------------------------------------------------------------------------------------------------------------------
	'Check the sprinter process if exist then terminate it 
	'--------------------------------------------------------------------------------------------------------------------------------
	  If StrComp(Process.Name, myProcess, vbTextCompare) = 0 Then 
	      Process.Terminate()                                   
	  End If
Next

If Browser(sVersion,"CreationTime:=0").Exist(3)  Then
	Reporter.ReportEvent micFail, "Close the all instances of Edge Chromium Browser", "Fail to Close the all instances of Edge Chromium Browser"					
	ExitComponent
Else
	Reporter.ReportEvent micPass, "Close the all instances of Edge Chromium Browser", "Successfully Close the all instances of Edge Chromium Browser"					
End If
'Next
'--------------------------------------------------------------------------------------------------------------------------------
'Select Instance Name
'--------------------------------------------------------------------------------------------------------------------------------
Select Case lCase(Parameter("str_Instance"))
	'--------------------------------------------------------------------------------------------------------------------------------
	Case "project"
		targetUrl = SearchAndLoadResourceByName("ActiveWorkspace2406_Data.xml").GetValue("PROJECT")
	'--------------------------------------------------------------------------------------------------------------------------------
	Case "prod"
		targetUrl = SearchAndLoadResourceByName("ActiveWorkspace_Data.xml").GetValue("PROD")
	'--------------------------------------------------------------------------------------------------------------------------------
	Case "prod-nonsso"
		targetUrl = SearchAndLoadResourceByName("ActiveWorkspace_Data.xml").GetValue("PROD-NonSSO")
	'--------------------------------------------------------------------------------------------------------------------------------
	Case "int2406" 
		targetUrl = SearchAndLoadResourceByName("ActiveWorkspace2406_Data.xml").GetValue("INT2406")
	'--------------------------------------------------------------------------------------------------------------------------------	
	Case "int2" 
		targetUrl = SearchAndLoadResourceByName("ActiveWorkspace_Data.xml").GetValue("INT2")
	'--------------------------------------------------------------------------------------------------------------------------------	
	Case "int1" 
		targetUrl = SearchAndLoadResourceByName("ActiveWorkspace_Data.xml").GetValue("INT1")
	'--------------------------------------------------------------------------------------------------------------------------------
	Case "pre"
		targetUrl = SearchAndLoadResourceByName("ActiveWorkspace_Data.xml").GetValue("PRE")
	'--------------------------------------------------------------------------------------------------------------------------------
	Case "pre-nonsso"
		targetUrl = SearchAndLoadResourceByName("ActiveWorkspace_Data.xml").GetValue("PRE-NONSSO")
	'--------------------------------------------------------------------------------------------------------------------------------
	Case "cloud"
		targetUrl = SearchAndLoadResourceByName("ActiveWorkspace_Data.xml").GetValue("CLOUD")
	'--------------------------------------------------------------------------------------------------------------------------------
	Case "int2-nonsso"
		targetUrl = SearchAndLoadResourceByName("ActiveWorkspace_Data.xml").GetValue("INT2-NONSSO")
	'--------------------------------------------------------------------------------------------------------------------------------
	Case Else 
			ExitComponent
	'--------------------------------------------------------------------------------------------------------------------------------			
End Select
'--------------------------------------------------------------------------------------------------------------------------------
'Invoke the Active Workspace (AWC) application 
'--------------------------------------------------------------------------------------------------------------------------------
'systemutil.Run "msedge.exe",targetUrl
InvokeApplication "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe "& targetUrl
Wait 5
objWshShell.SendKeys "{TAB}"
Wait 2
objWshShell.SendKeys "{TAB}"
Wait 2
objWshShell.SendKeys "{TAB}"
'If Browser("creationtime:=0").Exist Then
'    Browser("creationtime:=0").Sync
'     ' Optional: Refresh the browser if it appears blank
'	 Browser("creationtime:=0").Refresh
'End If
'

'Enter TAB to see the home page - 24 April 2024
'objWshShell.SendKeys "{TAB}"
'testing purpose
'Browser("Browser").Refresh
'wait 5
'For iCounter=0 to 5
'	If obj_AWCTeamcenterHome.exist Then
'		Exit For
'	Else
'		If Waituntilexist(obj_AWCTeamcenterHome, 4,6) = False  Then
'		       'Do Nothing
'		        Browser("Browser").Highlight
'		        wait 1
'	          	Browser("Browser").Refresh
'	          	wait 2
'		       objWshShell.SendKeys "{ENTER}"
'		       wait 1
'		       Browser("Browser").Refresh
'		       wait 1
'		       Browser("Browser").Sync
'		End If
'	End If
'Next
'wait 10

''TAB enter is working so commenting the below AI code - 24 -April 2024
'AIUtil.SetContext Browser("creationtime:=0")
'AIUtil.Context.SetBrowserScope(BrowserWindow)
'AIUtil.FindTextBlock("https:/,").Hover

'Non SSO is not valid any more 11 - oct -2023
If  lCase(Parameter("str_Instance"))="project"or lCase(Parameter("str_Instance"))="int2-nonsso" or lCase(Parameter("str_Instance"))="int2"  or  lCase(Parameter("str_Instance"))="cloud"  or  lCase(Parameter("str_Instance"))="int2406" Then
	'--------------------------------------------------------------------------------------------------------------------------------
	'Enter User Name
	'--------------------------------------------------------------------------------------------------------------------------------
	If Parameter("str_UserName")<>"" Then
		sUserName=Parameter("str_UserName")
	'	sUserName=GetResource("ActiveWorkspace_UserDetails.xml").GetValue(Parameter("str_UserName"))
		If sUserName="" Then
			Reporter.ReportEvent micFail, "Enter the Correct UserName", "Fail to Sign In to the Active Workspace as [ UserName is wrong ]"
			ExitComponent
		Else
		
			If WaitUntilExist(obj_AWCTeamcenterHome.WebEdit("wedit_UserName"), 2, 5)=False Then
			     objWshShell.SendKeys "{ENTER}"
			      Browser("Browser").Sync
			End  IF
			'Enter Username
			If Fn_Web_UI_WebEdit_Operations("Start Active Workspace","Set",obj_AWCTeamcenterHome, "wedit_UserName", sUserName ) Then
				Reporter.ReportEvent micPass, "Enter the [ Username ] in username edit field", "Successfully entered the [ "& sUserName &" ] in username edit field"
			Else
				Reporter.ReportEvent micFail, "Enter the [ Username ] in username edit field", "Fail to enter the [ Username ] in username edit field"
				ExitComponent
			End  If
			 Browser("Browser").Sync
		End  IF	
	Else
		Reporter.ReportEvent micFail, "Fail to Sign In to the Active Workspace", "Fail to Sign In to the Active Workspace as [ Username is Empty ]"
		ExitComponent
	End If
	'--------------------------------------------------------------------------------------------------------------------------------
	'Enter Password
	'--------------------------------------------------------------------------------------------------------------------------------
	If Parameter("str_UserName")<>"" Then
		sPassword=SearchAndLoadResourceByName("ActiveWorkspace2406_UserDetails.xml").GetValue(Parameter("str_UserName"))
		If sPassword="" Then
			Reporter.ReportEvent micFail, "Enter the Correct Password", "Fail to Sign In to the Active Workspace as [ Password is wrong ]"
			ExitComponent
		Else
			'Enter Password
			If Fn_Web_UI_WebEdit_Operations("Start Active Workspace","SetSecure",obj_AWCTeamcenterHome, "wedit_Password", sPassword ) Then
				Reporter.ReportEvent micPass, "Enter the [ Password ] in password edit field", "Successfully entered the [ " & sPassword & " ] in password edit field"
			Else
				Reporter.ReportEvent micFail, "Enter the [ Password ] in password edit field", "Fail to enter the [ Password ] in password edit field"
				ExitComponent
			End  If
			 Browser("Browser").Sync
		End  IF	
	Else
		Reporter.ReportEvent micFail, "Fail to Sign In to the Active Workspace", "Fail to Sign In to the Active Workspace as [ Password is Empty ]"
		ExitComponent
	End If
	'--------------------------------------------------------------------------------------------------------------------------------
	'Click on Login  button
	'--------------------------------------------------------------------------------------------------------------------------------	
	IF Fn_WEB_UI_WebButton_Operations("Start Active Workspace", "click", obj_AWCTeamcenterHome, "wbtn_Login","","","") Then
		Reporter.ReportEvent micPass, "Click on Log In button", "Successfully Clicked on Log In button"
	Else
		Reporter.ReportEvent micFail, "Click on Log In button", "Fail to Click on Log In button"
		ExitComponent
	End  If	
End If	

Call Fn_AWC_ReadyStatusSync(2)	
     
'''--------------------------------------------------------------------------------------------------------------------------------
'''Verify the header
'''--------------------------------------------------------------------------------------------------------------------------------
''''''For Production Execution
''''If Parameter("str_Location") = "CVS PLM (PRE-PROD)" Then
''''	Parameter("str_Location") = "CVS PLM (PRODUCTION)"
''''End If
''
''''For INT1 Execution
''''If Parameter("str_Location") = "CVS PLM (PRE-PROD)" Then
''''	Parameter("str_Location") = "CVS PLM (INT1)"
''''End If
''
'''objWshShell.SendKeys "{ENTER}"
'''Browser("Browser").Refresh
'''Call Fn_AWC_ReadyStatusSync(13)

Wait 2
objWshShell.SendKeys "{TAB}"
Wait 2
objWshShell.SendKeys "{TAB}"
''
If WaitUntilExist(obj_AWCTeamcenterHome, 2, 8)Then
'	Browser("Browser").Refresh
'	Call Fn_AWC_ReadyStatusSync(1)	
'	Parameter("str_Location") =Browser("Browser").GetROProperty("title")
'	sTemp =Split(Parameter("str_Location"),"-")
	For iCounter = 0 To 10
		If WaitUntilExist(obj_AWCTeamcenterHome.WebElement("wele_HomePageHeader"), 5,5) Then
			sHeaderText=obj_AWCTeamcenterHome.WebElement("wele_HomePageHeader").GetROProperty("outertext")
			Exit For
		Else
			Browser("Browser").Refresh
			Call Fn_AWC_ReadyStatusSync(1)	
		End  If
	Next
	
'	If Instr(trim(sHeaderText),trim(sTemp(0)))>0 Then
	If Instr(trim(sHeaderText),trim(Parameter("str_Location") ))>0 Then
		Reporter.ReportEvent micPass, "Verify the Header on Home page", "Successfully verified the Header [ "& trim(Parameter("str_Location") ) & " ] on the Home Page"
		Reporter.ReportEvent micPass, "SignIn to the Active  Workspace", "Successfully SignIn to the Active  Workspace "
	Else
		Reporter.ReportEvent micFail, "Verify the Header on Home page", "Fail to verify the Header [ "& trim(Parameter("str_Location") ) & " ] on the Home Page"
		Reporter.ReportEvent micFail, "SignIn to the Active  Workspace", "Fail to SignIn to the Active  Workspace"
		ExitComponent
	End If
Else
	Reporter.ReportEvent micFail, "SignIn to the Active  Workspace", "Fail to SignIn to the Active  Workspace"
	ExitComponent
End  If
'--------------------------------------------------------------------------------------------------------------------------------
'Store logged in user as output parameter
'--------------------------------------------------------------------------------------------------------------------------------
Call Fn_WEB_UI_WebButton_Operations("", "click", obj_AWCTeamcenterHome, "wbtn_UserProfile","","","")
Call Fn_AWC_ReadyStatusSync(1)
sTempUserNm=Fn_Web_UI_WebObject_Operations("Change your group, role,project,workspace and revision rule for your session", "getroproperty", obj_AWCTeamcenterHome.WebElement("wele_LoggedInUser"), "2", "title", "")
Parameter("str_UserName_Out")=Ucase(sTempUserNm)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
If Err.Number<> 0 Then
	Reporter.ReportEvent micFail, "Start Active Workspace", "Fail to perform [ Start Active Workspace ]  Operation due to [ " & Err.Description & " ]"
	ExitComponent
Else
	Reporter.ReportEvent micPass, "Start Active Workspace", "Successfully performed [ Start Active Workspace ] operation "
End If
'-----------------------------------------------------------------------------------------------------------------------------------------------------------	
'Set object nothing
Set obj_AWCTeamcenterHome=Nothing
Set objWshShell=Nothing
'-----------------------------------------------------------------------------------------------------------------------------------------------------------


