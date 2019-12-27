'All the common functions will be in this file

Dim strComputer
strComputer = "."


'Get the machine name
Public Function GetMachineName(WshShell, MachineName)
	Dim strmachine
	
	'Getting Machine name
	if MachineName="." Or Len(Trim(MachineName)) = 0 Then 
		Set objScriptExec = WshShell.Exec("ipconfig /all")
		strIpConfig = objScriptExec.StdOut.ReadAll
		arrStr=split(strIpConfig,vbNewLine)

		for i=lbound(arrstr) to ubound(arrstr)
			if (Rtrim(Ltrim(arrstr(i)))<>"") Then
				arrstrsub=split(arrstr(i),": ")
				If(UBound(arrstrsub) >= 1) Then
					if (UCASE(Left(Rtrim(LTRIM(arrstrsub(0))),9))="HOST NAME") then
						strmachine=Replace(Trim(arrstrsub(1))," ","")
						Exit for
					end If
				End If 
			end if
		Next
	else
		strmachine=Trim(MachineName)
	end If
	GetMachineName=strmachine
End Function

'Get the machine name
Public Function GetMachineNameSmart(MachineName)
	Dim strmachine
	
	'Getting Machine name
	if MachineName="." Or Len(Trim(MachineName)) = 0 Then 
		Set wshNetwork = WScript.CreateObject( "WScript.Network" )
		strmachine = wshNetwork.ComputerName
	else
		strmachine=Trim(MachineName)
        ' Hack, convert machine name to netbios name for names with len > 15
	    if Len(strmachine) > 15 then
		    strmachine = UCase(Left(strmachine, 15))
	    end if
	end If
	GetMachineNameSmart=strmachine
End Function

'Sets the environment variable PALEnvironment with the value from TopasDeploymentDetails.csv
Public Function SetPALEnvValue(WshShell, pwlog)
	Dim objusrfile, fs, strEnvironment
	
	SetPALEnvValue = false
	Set fs=CreateObject("Scripting.FileSystemObject")
	Set Envi = WshShell.Environment ("Process")
	strSystemDrive = Envi("SYSTEMDRIVE")
	Set objusrfile = fs.OpenTextFile(strSystemDrive +"\Topas\TopasDeploymentDetails.csv")
	
	'Get the environment name from TopasDeploymentDetails.csv
	Do while NOT objusrfile.AtEndOfStream
		strtmp=objusrfile.ReadLine
		arrStr1 = split(strtmp,",")
		If ubound(arrStr1) >= 1 Then
			If (StrComp("environment",arrStr1(0),1)=0) Then
				strEnvironment = Trim(arrStr1(1))
				Exit Do 
			End If
		End If
	Loop
	
	objusrfile.close
	set objusrfile=Nothing
	Set fs = nothing
	If Len(strEnvironment) = 0 Then
		pwlog.writeline Now & " : Environment name not found in TopasDeploymentDetails.csv"
		pwlog.close
		Wscript.quit(-1)
	Else
		pwlog.writeline Now & " : Environment name : '" & strEnvironment & "'"
	End If 
	
	SetPALEnvValue = SetEnvVariables(wshShell, "PALEnvironment", strEnvironment)
	
End Function

'Sets the environment variable for the current User
Public Function SetEnvVariables(wshShell, ByVal EnvVariable, ByVal EnvValue)
	'Set wshShell = CreateObject( "WScript.Shell" )
	'Set wshSystemEnv = wshShell.Environment( "SYSTEM" )
	Set wshUserEnv = wshShell.Environment( "USER" )
	' Set the environment variable
	wshUserEnv(EnvVariable) = EnvValue
	If(Err.Number = 0) Then 
		SetEnvVariables = True
	Else
		SetEnvVariables = False 
	End If 
	' Delete the environment variable
	'wshSystemEnv.Remove( "TestSystem" )
	Set wshSystemEnv = Nothing
	'Set wshShell     = Nothing
End Function

'Deletes the environment variable for the current User
Public Function DeleteEnvVariables(wshShell, ByVal EnvVariable)
	Set wshUserEnv = wshShell.Environment( "USER" )
	' Delete the environment variable
	wshUserEnv.Remove(EnvVariable)
	If(Err.Number = 0) Then
		DeleteEnvVariables = True
	Else
		DeleteEnvVariables = False 
	End If
	
	Set wshSystemEnv = Nothing
End Function

'Enable/Disable the windows service
'Returns 0 if sucessful, else it will return the error number
Public Function ModifyWindowsService(WshShell1, ByVal strServiceName, ByVal intEnable)
    Dim objWMIService, colServiceList, objService
    
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	' Get the service collection
	Set colServiceList = objWMIService.ExecQuery _
	("Select * from Win32_Service where name = '" & strServiceName &"'")

	For Each objService in colServiceList
		If 	intEnable > 0 Then
			Err.Number = objService.ChangeStartMode("Automatic")
			If Err.Number = 0 Then
				objService.StartService()
			End If
		Else
			Err.Number = objService.ChangeStartMode("Disabled")
			If Err.Number = 0 Then
				objService.StopService()
			End If
		End If
		If Err.Number <> 0 Then
			Exit For
		End If
	Next

	ModifyWindowsService = Err.Number
End Function


'Get the password for the username passed from the TopasDeploymentDetails.csv
Function GetPasswordFromCSV(strcon, pwlog)

	Dim strusername, strPassword, strnam, strTemp
	Dim objPasswordFile
	Dim objusrfile, fs, strEnvironment, WshShell

	Set WshShell = WScript.CreateObject("WScript.Shell")
	
	Set fs=CreateObject("Scripting.FileSystemObject")
	Set Envi = WshShell.Environment ("Process")
	strSystemDrive = Envi("SYSTEMDRIVE")

	If Not fs.fileexists(strSystemDrive +"\Topas\TopasDeploymentDetails.csv") Then
		pwlog.writeline Now & " :" +  strSystemDrive +"\Topas\TopasDeploymentDetails.csv file not found"
		pwlog.close
		wscript.quit(-1)
	End If

	Set objPasswordFile=fs.openTextFile(strSystemDrive +"\Topas\TopasDeploymentDetails.csv")
	Dim intUserNameColumn
	
	intUserNameColumn=0
	strnam=split(strcon,";")
	
	'Find the user name first
	'We may get the user name directly or connection string, the 1st case is for connection string
	If UBound(strnam) > 0 Then
		For j=0 to Ubound(strnam)
			if instr(lcase(strnam(j)),lcase("User Id")) > 0 Then
				strTemp = Split(strnam(j),"=")
				If UBound(strTemp) > 0 Then 
					strusername = Trim(strTemp(1))
				End If
				intUserNameColumn = 0
				Exit For
			End if
		Next
	Else ' this case is when user name is passed as paramenter
		strusername=strcon
		intUserNameColumn = 1
	End If 
	
	If Len(strusername) > 0 Then 
		'Get the password for coressponding user
		Do while NOT objPasswordFile.AtEndOfStream

			arrStr1 = split(objPasswordFile.ReadLine,",")
			If UBound(arrStr1) >= intUserNameColumn Then
				'if (StrComp(ucase(ltrim(rtrim(strusername))),ucase(ltrim(rtrim(arrStr1(0)))),1)=0) and (ucase(arrStr1(2))="DB") then
				if (StrComp(ucase(strusername),ucase(arrStr1(intUserNameColumn)),1)=0) then
					strPassword = arrStr1(intUserNameColumn+1)	
					Exit DO
				end If
			End If 
		Loop

		objPasswordFile.Close
		Set objPasswordFile = Nothing 
	End If 

	If strPassword = "" Then 
		pwlog.writeline Now & " : Password not found in TopasDeploymentDetails.csv for the user : " & strusername
		pwlog.close
		wscript.quit(-1)
	Else
		pwlog.writeline Now & " : Found password in TopasDeploymentDetails.csv for the user : " & strusername
	End If 

	GetPasswordFromCSV = strPassword

End Function



'sample function
Public Function SayHello()
	MsgBox "Hi"
End Function 

'Call Test()

Function Test()
	Set WshShell = WScript.CreateObject("WScript.Shell")
	'msgbox ModifyWindowsService(WshShell,"SR_Service",1)
	Set fs=CreateObject("Scripting.FileSystemObject")
	Set pwlog=fs.OpenTextFile("c:\Temp\ASAPConfiguration.log",8, True)
	MsgBox GetPasswordFromCSV("Data Source=US1S;User ID = PALHEALTHMONITOR;Password=devpalhealthmonitor;Enlist=false",pwlog)
End Function