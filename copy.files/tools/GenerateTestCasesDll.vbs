Dim fs 
Set fs=CreateObject("Scripting.FileSystemObject")


Function GeneratePWAssemblyListFile()
	Dim nunitlist
	Set nunitlist = CreateObject("System.Collections.ArrayList")

	nunitlist.Add "AssetIdentifierServices.dll"
	nunitlist.Add "CurrencyServices.dll"
	nunitlist.Add "PWServiceImplementation.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Apps.TopasEnterprise.TopasEnterpriseServer.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Common.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Services.AAALiveServices.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Services.AttributeServices.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Services.CorporateAction.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Services.OrganizationServices.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Services.PerformanceServices.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Services.PortfolioServices.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Services.ProfilingServices.dll"
	nunitlist.Add "TimeSeries.dll"

	Call GenerateListFile(".\src\Financial\Vestek\Apps\TopasEnterprise\PortfolioWarehouse\PWWebService\bin", nunitlist, "tools\UnitTestAssemblies\PW\UnitTestAssemblies.txt")
End Function 


Function GeneratePAAssemblyListFile()
	Dim nunitlist
	Set nunitlist = CreateObject("System.Collections.ArrayList")
	
	nunitlist.Add "AssetIdentifierServices.dll"
	nunitlist.Add "CurrencyServices.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Apps.TopasEnterprise.TopasEnterpriseServer.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Common.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Services.AAALiveServices.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Services.AttributeServices.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Services.CorporateAction.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Services.OrganizationServices.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Services.PerformanceServices.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Services.PortfolioServices.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Services.ProfilingServices.dll"
	nunitlist.Add "TimeSeries.dll"
	nunitlist.Add "TPADataTransformation.dll"
	nunitlist.Add "TPAServiceImplementations.dll"
	' nunitlist.Add "AnalyticsApps.Tests.dll"

	Call GenerateListFile(".\src\Financial\Vestek\Apps\TopasEnterprise\PortfolioAnalytics\TPAWebServices\bin", nunitlist, "tools\UnitTestAssemblies\PA\UnitTestAssemblies.txt")
End Function 


Function GenerateASAPAssemblyListFile()
	Dim nunitlist
	Set nunitlist = CreateObject("System.Collections.ArrayList")

	nunitlist.Add "AssetIdentifierServices.dll"
	nunitlist.Add "CurrencyServices.dll"
	nunitlist.Add "JobManager.dll"
	nunitlist.Add "PortfolioLoadJobMgmt_UnitTests.dll"
	nunitlist.Add "RiskUtil.dll"
	nunitlist.Add "SecurityLoadJobMgmt.dll"
	nunitlist.Add "StatisticalFactorEngine.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Apps.TopasEnterprise.TopasEnterpriseServer.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Common.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Services.AAALiveServices.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Services.AttributeServices.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Services.CorporateAction.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Services.OrganizationServices.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Services.PerformanceServices.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Services.PortfolioServices.dll"
	nunitlist.Add "Thomson.Financial.Vestek.Services.ProfilingServices.dll"
	nunitlist.Add "TimeSeries.dll"
	nunitlist.Add "TransformService.dll"

	Call GenerateListFile(".\src\Financial\Vestek\Apps\JobManagement\ASAPWCFService\ASAPService\bin", nunitlist, "tools\UnitTestAssemblies\ASAP\UnitTestAssemblies.txt")
End Function 


Function GenerateListFile(strPath, List, outputFileName )
	Set ListFile 	= fs.CreateTextFile(outputFileName)
	
	For Each dll in List
		ListFile.Write  strPath + "\" + dll + " "
	Next   

	ListFile.close
	Set ListFile = Nothing
End Function 





