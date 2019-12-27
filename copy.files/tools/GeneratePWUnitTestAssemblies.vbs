
On Error Resume Next 

Call IncludeFileCommon()

Call GeneratePWAssemblyListFile()




'This function will include the common.vbs file tpmconfig folder
Function IncludeFileCommon()
	Call llf_filevbscriptload("GenerateTestCasesDll.vbs")
End Function 

'call this function to use the vbscripts fucntions in some other file
'call this function to use the vbscripts fucntions in some other file
Public Function llf_filevbscriptload(byval rstrfilename)
	Const forreading = 1 
	Dim objfilesystem 
	Dim objfile, strBasePath 
	Set objfilesystem = createobject("scripting.filesystemobject") 
	
	strBasePath = objfilesystem.GetFolder(".").Path 
	strBasePath = Replace(strBasePath,"\tools","")
	rstrfilename = strBasePath & "\tools\" & rstrfilename

	If objfilesystem.fileexists(rstrfilename) = True Then
		Set objfile = objfilesystem.opentextfile(rstrfilename, forreading, False)
		llf_filevbscriptload = objfile.readall
		objfile.close
		Set objfile = Nothing 
		Call executeglobal(llf_filevbscriptload)
	Else 
		wscript.quit(-1)
	End if 
	
	set objfilesystem = nothing 
 End Function  