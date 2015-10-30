if(!(Test-Path scripts:)) { New-PSDrive -Name scripts -PSProvider FileSystem -Root "E:\Prod_E\0AP\Scripts" }
. scripts:\Get-FileName.ps1
. scripts:\Execute-CommandManager.ps1

Function Get-ProjectConfiguration {
<#
		.NAME
		 Get-ProjectConfiguration
		.DESCRIPTION
         Gets the project configuration for a MicroStrategy project.
        .NOTES
		 		 			
         AUTHOR....:  Craig Grady
         LAST EDIT.:  01/21/2015
         CREATED...:  01/21/2015
#>
       
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[string] $ProjectName,
[string] $FileName
)
$Empty = [String]::Empty

$CmdLine = "LIST ALL PROPERTIES FOR PROJECT CONFIGURATION IN PROJECT `"$ProjectName`";`r`n"
if($FileName -ne $Empty) { Out-File -FilePath $FileName -Encoding ascii -InputObject $CmdLine } else { $CmdLine }
}

Function Get-ProjectConfigurationXML {
<#
		.NAME
		 Get-ProjectConfigurationXML
		.DESCRIPTION
         Gets the project configuration for a MicroStrategy project that is saved as an XML file.
        .NOTES
		 		 			
         AUTHOR....:  Craig Grady
         LAST EDIT.:  02/01/2015
         CREATED...:  01/21/2015
		
		.EXAMPLES
		$pc = Get-ProjectConfigurationXML -UseFileDialog
		
		$pc = Get-ProjectConfigurationXML -FilePath tmp.xml
#>

[CmdletBinding()]
param(
[Parameter(ParameterSetName='File', Mandatory=$True)]
[string] $FilePath,
[Parameter(ParameterSetName='Dialog', Mandatory=$True)]
[Switch] $UseFileDialog
)

if($UseFileDialog) { $FilePath = Get-FileName($env:UserProfile) }

[xml]$getconfig = Get-Content $FilePath
 $getconfig.CommandManagerResults.ListPropertiesProjectConfiguration.Row
}

Function Get-Element-Advanced {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Advanced
}

Function Get-Element-Caching {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Caching
}

Function Get-Element-DBInstances {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.DBInstances
}

Function Get-Element-Deliveries {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Deliveries
}

Function Get-Element-DocumentsAndReports {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.DocumentsAndReports
}

Function Get-Element-Drillings {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Drillings
}

Function Get-Element-ExportSettings {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ExportSettings
}

Function Get-Element-Governings {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings
}

Function Get-Element-IntelligentCubes {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.IntelligentCubes
}

Function Get-Element-ObjectTemplates {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ObjectTemplates
}

Function Get-Element-PDFSettings {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.PDFSettings
}

Function Get-Element-ProjectAccesses {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ProjectAccesses
}

Function Get-Element-Governings-MaxFileSizeImport {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.MaxFileSizeImport
}

Function Get-Element-Governings-MaxQuotaImport {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.MaxQuotaImport
}

Function Get-Element-Governings-MaxFileSubscriptions {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.MaxFileSubscriptions
}

Function Get-Element-PDFSettings-PDFHeaderRight {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.PDFSettings.PDFHeaderRight
}

Function Get-Element-ReportDefinition-GraphResultSet {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ReportDefinition.GraphResultSet
}

Function Get-Element-Governings-MaxUserSessionsPerProject {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.MaxUserSessionsPerProject
}

Function Get-Element-Deliveries-EnablePrintRange {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Deliveries.EnablePrintRange
}

Function Get-Element-Governings-MaxPrintSubscriptions {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.MaxPrintSubscriptions
}

Function Get-Element-IntelligentCubes-IntelligentCubeFileDir {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.IntelligentCubes.IntelligentCubeFileDir
}

Function Get-Element-DocumentsAndReports-ReportDetails {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.DocumentsAndReports.ReportDetails
}

Function Get-Element-Advanced-DocumentDirectory {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Advanced.DocumentDirectory
}

Function Get-Element-Governings-MaxElementRows {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.MaxElementRows
}

Function Get-Element-Drillings-DrillPersonalizedPath {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Drillings.DrillPersonalizedPath
}

Function Get-Element-Governings-MaxHLSubscriptions {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.MaxHLSubscriptions
}

Function Get-Element-Caching-DoNotCreateOrUpdateMatchingCaches {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Caching.DoNotCreateOrUpdateMatchingCaches
}

Function Get-Element-ObjectTemplates-TemplateTPL {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ObjectTemplates.TemplateTPL
}

Function Get-Element-ProjectAccesses-SecurityRole {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ProjectAccesses.SecurityRole
}

Function Get-Element-RightToLeft {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.RightToLeft
}

Function Get-Element-Governings-WaitTimeForPromptAnswers {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.WaitTimeForPromptAnswers
}

Function Get-Element-Governings-MaxJobsPerUserAccount {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.MaxJobsPerUserAccount
}

Function Get-Element-ExportSettings-ExportHeader {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ExportSettings.ExportHeader
}

Function Get-Element-ProjectStatuses-ShowStatus {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ProjectStatuses.ShowStatus
}

Function Get-Element-ObjectTemplates-UseDocumentWizard {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ObjectTemplates.UseDocumentWizard
}

Function Get-Element-IntelligentCubes-CreateIntelligentCubesByDatabaseConnection {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.IntelligentCubes.CreateIntelligentCubesByDatabaseConnection
}

Function Get-Element-IntelligentCubes-IntelligentCubeMaxRam {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.IntelligentCubes.IntelligentCubeMaxRam
}

Function Get-Element-DBInstances-DBInstance {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.DBInstances.DBInstance
}

Function Get-Element-ObjectTemplates-MetricTPL {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ObjectTemplates.MetricTPL
}

Function Get-Element-Governings-MaxMobileSubscriptions {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.MaxMobileSubscriptions
}

Function Get-Element-Governings-MemoryConsumptionDuringSQLGeneration {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.MemoryConsumptionDuringSQLGeneration
}

Function Get-Element-Governings-WarehouseExecutionTime {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.WarehouseExecutionTime
}

Function Get-Element-ObjectTemplates-MetricShowEmptyTPL {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ObjectTemplates.MetricShowEmptyTPL
}

Function Get-Element-Caching-RerunFileEmailOrPrintSubscriptionsAgainstTheWarehouse {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Caching.RerunFileEmailOrPrintSubscriptionsAgainstTheWarehouse
}

Function Get-Element-Advanced-MaxAttributeElements {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Advanced.MaxAttributeElements
}

Function Get-Element-Drillings-DrillToImmediate {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Drillings.DrillToImmediate
}

Function Get-Element-Advanced-CustomPromptStyles {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Advanced.CustomPromptStyles
}

Function Get-Element-IntelligentCubes-IntelligentCubeMaxNo {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.IntelligentCubes.IntelligentCubeMaxNo
}

Function Get-Element-Project {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Project
}

Function Get-Element-SecurityFilterResultSet-SecurityFilterMerge {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.SecurityFilterResultSet.SecurityFilterMerge
}

Function Get-Element-Deliveries-DeliverSubscriptionWithNoData {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Deliveries.DeliverSubscriptionWithNoData
}

Function Get-Element-Governings-MaxScheduleReportExecTime {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.MaxScheduleReportExecTime
}

Function Get-Element-Deliveries-AppendFooter {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Deliveries.AppendFooter
}

Function Get-Element-IntelligentCubes-LoadIntelligentCubesOnStartup {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.IntelligentCubes.LoadIntelligentCubesOnStartup
}

Function Get-Element-ObjectTemplates-DocumentTPL {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ObjectTemplates.DocumentTPL
}

Function Get-Element-PDFSettings-PDFHeaderCenter {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.PDFSettings.PDFHeaderCenter
}

Function Get-Element-Deliveries-DeliverSubscriptionWithPartialResults {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Deliveries.DeliverSubscriptionWithPartialResults
}

Function Get-Element-ObjectTemplates-TemplateShowEmptyTPL {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ObjectTemplates.TemplateShowEmptyTPL
}

Function Get-Element-Governings-InteractiveJobsPerProject {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.InteractiveJobsPerProject
}

Function Get-Element-ReportDefinition-NullValues {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ReportDefinition.NullValues
}

Function Get-Element-Advanced-DisplayAttributeHierarchyPrompt {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Advanced.DisplayAttributeHierarchyPrompt
}

Function Get-Element-Governings-MaxReportExecutionTime {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.MaxReportExecutionTime
}

Function Get-Element-Caching-KeepDocumentAvailableForManipulationForHistoryListSubscriptionsOnly {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Caching.KeepDocumentAvailableForManipulationForHistoryListSubscriptionsOnly
}

Function Get-Element-ProjectStatuses-ProjectStatus {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ProjectStatuses.ProjectStatus
}

Function Get-Element-Governings-MaxIntelligentCubeResultRows {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.MaxIntelligentCubeResultRows
}

Function Get-Element-ProjectDescription {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Description
}

Function Get-Element-Deliveries-EmailFooter {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Deliveries.EmailFooter
}

Function Get-Element-ExportSettings-ExportFlashFormat {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ExportSettings.ExportFlashFormat
}

Function Get-Element-Caching-RerunHistoryListAndMobileSubscriptionsAgainstTheWarehouse {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Caching.RerunHistoryListAndMobileSubscriptionsAgainstTheWarehouse
}

Function Get-Element-Governings-MaxJobsPerProject {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.MaxJobsPerProject
}

Function Get-Element-SchedulesResultSet-Schedule {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.SchedulesResultSet.Schedule
}

Function Get-Element-Caching-CacheEncryptionLevelOnDisk {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Caching.CacheEncryptionLevelOnDisk
}

Function Get-Element-Drillings-Drilling {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Drillings.Drilling
}

Function Get-Element-Advanced-EnablePersonalAnswers {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Advanced.EnablePersonalAnswers
}

Function Get-Element-SecurityModule-FreeformSQLAndMDXSecurity {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.SecurityModule.FreeformSQLAndMDXSecurity
}

Function Get-Element-ProjectStatuses-StatusOnTop {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ProjectStatuses.StatusOnTop
}

Function Get-Element-PDFSettings-PDFFooterRight {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.PDFSettings.PDFFooterRight
}

Function Get-Element-Governings-MaxReportResultRows {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.MaxReportResultRows
}

Function Get-Element-DocumentsAndReports-Watermark {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.DocumentsAndReports.Watermark
}

Function Get-Element-PDFSettings-PDFHeaderLeft {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.PDFSettings.PDFHeaderLeft
}

Function Get-Element-ExportSettings-DoNotMergeOrDuplicateHeadersWhenExportingToExcelWord {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ExportSettings.DoNotMergeOrDuplicateHeadersWhenExportingToExcelWord
}

Function Get-Element-PDFSettings-PDFFooterLeft {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.PDFSettings.PDFFooterLeft
}

Function Get-Element-ExportSettings-ExportFooter {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ExportSettings.ExportFooter
}

Function Get-Element-Advanced-EnableObjectDeletion {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Advanced.EnableObjectDeletion
}

Function Get-Element-Governings-MaxResultRows {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.MaxResultRows
}

Function Get-Element-Governings-MaxJobsPerUser {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.MaxJobsPerUser
}

Function Get-Element-DBInstances-DefaultDatamart {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.DBInstances.DefaultDatamart
}

Function Get-Element-IntelligentCubes-LoadIntelligentCubesIntoIntelligentServerMemoryUponPublication {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.IntelligentCubes.LoadIntelligentCubesIntoIntelligentServerMemoryUponPublication
}

Function Get-Element-Drillings-DrillSortOption {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Drillings.DrillSortOption
}

Function Get-Element-ProjectAccesses-Members {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ProjectAccesses.Members
}

Function Get-Element-CreateProfileAtLogin {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.CreateProfileAtLogin
}

Function Get-Element-ObjectTemplates-ReportShowEmptyTPL {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ObjectTemplates.ReportShowEmptyTPL
}

Function Get-Element-Governings-MaxDataMartResultRows {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.MaxDataMartResultRows
}

Function Get-Element-ReportDefinition-Advance {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ReportDefinition.Advance
}

Function Get-Element-ReportDefinition-SQLGeneration {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ReportDefinition.SQLGeneration
}

Function Get-Element-Governings-MaxCacheUpdateSubscriptions {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.MaxCacheUpdateSubscriptions
}

Function Get-Element-Caching-DoNotApplyAutomaticExpirationLogicForReportsContainingDynamicDates {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Caching.DoNotApplyAutomaticExpirationLogicForReportsContainingDynamicDates
}

Function Get-Element-DocumentsAndReports-WebServerMacro {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.DocumentsAndReports.WebServerMacro
}

Function Get-Element-ObjectTemplates-ReportTPL {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ObjectTemplates.ReportTPL
}

Function Get-Element-Governings-MaxEmailSubcriptions {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.MaxEmailSubcriptions
}

Function Get-Element-ObjectTemplates-DocumentShowEmptyTPL {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ObjectTemplates.DocumentShowEmptyTPL
}

Function Get-Element-DBInstances-UseWHLoginExecution {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.DBInstances.UseWHLoginExecution
}

Function Get-Element-Governings-MaxJobsPerUserSession {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.MaxJobsPerUserSession
}

Function Get-Element-PDFSettings-PDFFooterCenter {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.PDFSettings.PDFFooterCenter
}

Function Get-Element-ExportSettings-RunExportToExcelWordIn81CompabilityMode {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.ExportSettings.RunExportToExcelWordIn81CompabilityMode
}

Function Get-Element-DocumentsAndReports-EnableLinksInExportedFlashDocuments {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.DocumentsAndReports.EnableLinksInExportedFlashDocuments
}

Function Get-Element-Governings-MaxConcurrentInteractiveUserSessionsPerUser {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.Governings.MaxConcurrentInteractiveUserSessionsPerUser
}

Function Get-Element-SecurityModule-ProjectDefinitionSecurity {
[CmdletBinding()]
param(
[Parameter(Mandatory=$True)]
[System.Xml.XmlElement] $XMLFileByReference
)
$XMLFileByReference.SecurityModule.ProjectDefinitionSecurity
}


Function Set-ProjectConfiguration {
<#
		.NAME
		 Set-ProjectConfiguration
		.DESCRIPTION
         Sets the project configuration for a MicroStrategy project.
        .NOTES
		 		 			
         AUTHOR....:  Craig Grady
         LAST EDIT.:  01/21/2015
         CREATED...:  01/21/2015
		
		
		MicroStrategy Command Manager Alter project configuration syntax:
		ALTER PROJECT CONFIGURATION [DESCRIPTION "<Project_description>"] [WAREHOUSE "<WH_name>"] [STATUS "<html_input_file>"] [SHOWSTATUS (TRUE | FALSE)] [STATUSONTOP (TRUE | FALSE)] [DOCDIRECTORY "<folder_path>"] [MAXNOATTRELEMS <no_attribute_elems>] [USEWHLOGINEXEC (TRUE | FALSE)] [ENABLEOBJECTDELETION (TRUE | FALSE)] [MAXREPORTEXECTIME <no_seconds>] [MAXSCHEDULEREPORTEXECTIME <no_seconds>] [MAXNOINTELLIGENTCUBERESULTROWS <no_rows>] [MAXNODATAMARTRESULTROWS <no_rows>] [MAXNOREPORTRESULTROWS <no_rows>] [MAXNOELEMROWS <no_rows>] [MAXNOINTRESULTROWS <no_rows>] [MAXJOBSUSERACCT <no_jobs>] [MAXJOBSUSERSESSION <no_jobs>] [MAXEXECJOBSUSER <no_jobs>] [MAXJOBSPROJECT <no_jobs>] [MAXUSERSESSIONSPROJECT <no_user_sessions>] [MAXINTERACTIVESESSIONSPERUSER <no_concurrent_sessions>] [MAXHISTLISTSUBSCRIPTIONS <no_histlist_subscriptions>] [MAXCACHEUPDATESUBSCRIPTIONS <no_cacheupdate_subscriptions>] [MAXEMAILSUBSCRIPTIONS <no_email_subscriptions>] [MAXFILESUBSCRIPTIONS <no_file_subscriptions>] [MAXPRINTSUBSCRIPTIONS <no_print_subscriptions>] [MAXMOBILESUBSCRIPTIONS <no_mobile_subscriptions>] [MAXFILESIZEIMPORT <numbre_of_MB>] [MAXQUOTAIMPORT <number_of_MB>] [PROJDRILLMAP "<drill_map>" [IN FOLDER "<drill_map_location_path>"]] [REPORTTPL "<report_template>"] [REPORTSHOWEMPTYTPL (TRUE | FALSE)] [TEMPLATETPL "<template_template>"] [TEMPLATESHOWEMPTYTPL (TRUE | FALSE)] [METRICTPL "<metric_template>"] [METRICSHOWEMPTYTPL (TRUE | FALSE)] [DOCTPL "<document_template>"] [DOCSHOWEMPTYTPL (TRUE | FALSE)] [USEDOCWIZARD (TRUE | FALSE)] [INTELLIGENTCUBEFILEDIR "<folder_path>"] [INTELLIGENTCUBEMAXRAM <number_of_MB>] [INTELLIGENTCUBEMAXNO <no_cubes>] [DRILLIMMEDIATE (TRUE | FALSE)] [SORTDRILLPATHS (TRUE | FALSE)] [PDFHEADERLEFT "<header_text>"]  [PDFHEADERCENTER "<header_text>"]  [PDFHEADERRIGHT "<header_text>"] [PDFFOOTERLEFT "<header_text>"] [PDFFOOTERCENTER "<header_text>"] [PDFFOOTERRIGHT "<header_text>"] [EXPORTHEADER "<export_header>"]  [EXPORTFOOTER "<export_footer>"] [EXPORTCOMPATIBILITYMODE (TRUE|FALSE)] [MERGEHEADERS (TRUE|FALSE)] [EXPORTFLASHFORMAT (PDF|MHT)] [CREATEUSERPROFILEATLOGIN (TRUE|FALSE)] [REPORTDESCRIPTION (TRUE|FALSE)] [PROMPTDETAILS (TRUE|FALSE)] [FILTERDETAILS (TRUE|FALSE)] [TEMPLATEDETAILS (TRUE|FALSE)] [INCLUDEPROMPTTITLE  (TRUE|FALSE)] [REPLACEUNANSWEREDPROMPT (DEFAULT|BLANK|PROMPTNOTANSWERED|NOSELECTION|ALLNONE)] [SHOWATTRIBUTENAMEINPROMPT (NO|YES|REPEATED)] [INCLUDEUNUSEPROMPTS (TRUE|FALSE)] [USEDELIMITERSINREPORTOBJECTS (NO|YES|AUTOMATIC)] [USEALIASESINFILTERSDETAILS (TRUE|FALSE)] [REPORTFILTER (TRUE|FALSE)] [REPORTFILTERNAME (NO|YES|AUTOMATIC)] [REPORTFILTERDESC (TRUE|FALSE)] [REPORTLIMITS (TRUE|FALSE)] [VIEWFILTER (TRUE|FALSE)] [METRICQUALIFICATIONVIEWFILTER (TRUE|FALSE)] [DRILLFILTER (TRUE|FALSE)] [SECURITYFILTER (TRUE|FALSE)] [INCLUDEFILTERTYPE (TRUE|FALSE)] [SHOWEMPTYEXPRESSION (TRUE|FALSE)] [NEWLINEAFTERTYPENAME (TRUE|FALSE)] [NEWLINEBETWEENFILTER (TRUE|FALSE)] [SHOWREPORTLIMITS (BEFOREVIEWFILTER|AFTERVIEWFILTER)] [EXPANDSHORCUTFILTERS (SHOWDEFINITION | SHOWNAME  | SHOWNAMEANDDEFINITION)] [SHOWATTRIBUTEINLISTCONDITIONS (NO | YES | REPEATED)] [SEPARATORAFTERATTRNAME "<separator>"] [NEWLINEAFTERATTRNAME (TRUE|FALSE)] [SEPARATORBETWEENLASTTWOELEMENTS (OR | AND | COMMA | CUSTOM "<separator>")] [NEWLINEBETWEENELEMENTS (TRUE|FALSE)] [TRIMELEMENTS (TRUE|FALSE)] [DISPLAYOPERATORS (SYMBOLS|NAMES)] [INCLUDEATTRFORMINCONDITION (TRUE|FALSE)] [DYNAMICDATE (DEFAULT | SHOWDATE | SHOWEXPRESSIONS)] [NEWLINEBETWEENCONDITIONS (NO | YES | AUTOMATIC)] [LINESPACING (SINGLESPACE | DOUBLESPACE)] [PARENTHESESAROUNDCONDITIONS (NO | YES | AUTOMATIC)] [LOGICALOPBETWEENCONDITIONS (NO | YES | AND | OR)] [UNITSFROM (VIEW | BASE)] [BASETEMPLATE (NO | YES | AUTOMATIC)] [TEMPLATEDESCRIPTION (TRUE|FALSE)] [NONMETRICTEMPLATEUNITS (TRUE|FALSE)] [METRICS (TRUE|FALSE)] [ONLYCONDITIONAL (TRUE|FALSE)] [FORMULA (TRUE|FALSE)] [DIMENSIONALITY (TRUE|FALSE)] [CONDITIONALITY (TRUE|FALSE)] [TRANSFORMATION (TRUE|FALSE)] [(NOWATERMARK | TEXTWATERMARK TEXT "<text>" SIZEAUTOMATICALLY (TRUE|FALSE) WASHOUT (TRUE|FALSE) ORIENTATION (DIAGONAL|HORIZONTAL) | IMAGEWATERMARK SOURCE "<image_source>" SCALE (<scale_percentage> | AUTO))] [ALLOWDOCUMENTOVERWRITEWATERMARK (TRUE|FALSE)] [WEBSERVERMACRO "<webserver_address>"] [ENABLELINKSINMHT (TRUE|FALSE)] [CUSTOMPROMPSTYLES (VALUEPROMPT | OBJECTPROMPT | ATTRIBUTEPROMPT | HIERARCHYPROMPT | ATTRIBUTEQUALPROMPT | METRICQUALIFICATION) NAME "<prompt_name>" DESCRIPTION "<prompt_description>" BASESTYLE "<base_style>" [, (VALUEPROMPT | OBJECTPROMPT | ATTRIBUTEPROMPT | HIERARCHYPROMPT | ATTRIBUTEQUALPROMPT | METRICQUALIFICATION) NAME "<prompt_name>" DESCRIPTION "<prompt_description>" BASESTYLE "<base_style>" ...]] [DISPLAYATTRINALPHABETICALORDER (TRUE|FALSE)] [ENABLEPERSONALANSWERS (TRUE|FALSE)] [EXPORTPDFINHEBREW (TRUE|FALSE)] [DEFAULTDATAMART "<db_instance_name>"] [WAITTIMEFORPROMPTANSWERS <wait_time>] [WAREHOUSEEXECUTIONTIME <max_execution_time>] [SQLGENERATIONMEMORY <MB_for_SQL_generation>] [INTERACTIVEJOBPERPROJECT <no_interactive_jobs>] [CACHEENCRYPTION (NONE|LOW|HIGH)] [DISABLEAUTOMATICEXPIRATIONDYNAMICDATES (TRUE|FALSE)] [RERUNMOBILESUBSCRIPTIONS (TRUE|FALSE)] [RERUNEMAILSUBSCRIPTIONS (TRUE|FALSE)] [DISABLECREATINGMATCHINGCACHES (TRUE|FALSE)] [KEEPDOCUMENTFORMANIPULATION (TRUE|FALSE)] [CREATECUBESBYDATABASECONNECTION (TRUE|FALSE)] [LOADCUBESONSTARTUP (TRUE|FALSE)] [LOADCUBESONPUBLICATION (TRUE|FALSE)] [SECURITYFILTERMERGER (UNION|INTERSECT)] [ATTRSHOWNUMERIC (TRUE|FALSE)] [ATTRSHOWCHARACTER (TRUE|FALSE)] [ATTRSHOWBINARY (TRUE|FALSE)] [ATTRSHOWDATE (TRUE|FALSE)] [ATTRSHOWBIGDECIMAL (TRUE|FALSE)] [ATTRREPLACEUNDERSCORE (TRUE|FALSE)] [ATTRREMOVEID (TRUE|FALSE)] [ATTRFIRSTLETTERUPPERCASE (TRUE|FALSE)] [ATTRIDSTRING "<prefix>"] [ATTRDESCSTRING "<prefix>"] [ATTRLOOKUPSTRING "<prefix>"] [FACTSHOWNUMERIC (TRUE|FALSE)] [FACTSHOWCHARACTER (TRUE|FALSE)] [FACTSHOWBINARY (TRUE|FALSE)] [FACTSHOWDATE (TRUE|FALSE)] [FACTSHOWBIGDECIMAL (TRUE|FALSE)] [FACTREPLACEUNDERSCORE (TRUE|FALSE)] [FACTFIRSTLETTERUPPERCASE (TRUE|FALSE)] [VERIFYCUSTOMCOLUMNNAME (TRUE|FALSE)] [NULLDISPLAYWAREHOUSE "<string_to_display>"] [NULLDISPLAYCROSSTABULATOR "<string_to_display>"] [REPLACENULLWHENSORTED (TRUE|FALSE)] [REPLACENULLWHENSORTEDVALUE <replace_value>] [NOTCALCULATEDDISPLAY "<string_to_display>"] [MISSINGDISPLAY "<string_to_display>"] [OVERRIDEGRAPHTEMPLATEFONT (TRUE|FALSE)]  [CHARACTERSET "<character_set>"] [FONT "<font_name>"] [USEZEROFORGRAPHNULL (TRUE|FALSE)] [GRAPHROUNDEDEFFECT (NONE|STANDARD|OPTIMIZE)] [EMPTYREPORTMESSAGE "<message>"] [DISPLAYEMPTYMESSAGEINDOCUMENT (TRUE|FALSE)] [KEEPPAGEBYWHENSAVING (TRUE|FALSE)] [OVERWRITEWITHOLAPREPORTS (ALLOW|ALLOWWITHWARNING|PREVENT)] [MOVESORTKEYSWITHPIVOT (TRUE|FALSE)] [APPENDEMAILFOOTER (TRUE|FALSE)] [EMAILFOOTERTEXT "<footer_text>"] [ENABLEPRINTINGRANGE (TRUE|FALSE)] [DELIVERSUBSCRIPTIONWITHNODATA (TRUE|FALSE)] [DELIVERSUBSCRIPTIONWITHPARTIALRESULTS (TRUE|FALSE)] [(ALLSCHEDULES | SCHEDULES "<schedule_name>" [,"<schedule_name>" ...])] IN PROJECT "<project_name>";

		

		ALTER PROJECT CONFIGURATION DESCRIPTION "MicroStrategy Tutorial Project (version 7.2)" 
		WAREHOUSE "Tutorial Data" DOCDIRECTORY "C:\Program Files\MicroStrategy\Tutorial Reporting" 
		MAXNOATTRELEMS 1000 USEWHLOGINEXEC FALSE MAXREPORTEXECTIME 600 MAXNOREPORTRESULTROWS 
		32000 MAXNOELEMROWS 32000 MAXNOINTRESULTROWS 32000 MAXJOBSUSERACCT 100 
		MAXJOBSUSERSESSION 100 MAXEXECJOBSUSER 100 MAXJOBSPROJECT 1000 MAXUSERSESSIONSPROJECT 
		500 IN PROJECT "MicroStrategy Tutorial";

		ALTER PROJECT CONFIGURATION PROJDRILLMAP 
		"Tutorial Standard Drill Map" IN FOLDER "\Public Objects\Drill Maps" REPORTTPL "Blank Report"
		REPORTSHOWEMPTYTPL FALSE TEMPLATESHOWEMPTYTPL FALSE
		METRICSHOWEMPTYTPL FALSE IN PROJECT "MicroStrategy Tutorial";

		LIST ALL PROPERTIES FOR PROJECT CONFIGURATION IN PROJECT "MicroStrategy Tutorial";

		ALTER STATISTICS DBINSTANCE "Tutorial Data" BASICSTATS ENABLED DETAILEDREPJOBS TRUE JOBSQL TRUE 
		IN PROJECT "MicroStrategy Tutorial";

		REMOVE DBINSTANCE "Excel Data Source" FROM PROJECT "MicroStrategy Tutorial";

		ADD DBINSTANCE "Excel Data Source" IN PROJECT "MicroStrategy Tutorial";

		ALTER PROJECT CONFIGURATION STATUS "\\Production_Machine\StatusFiles\Status.txt" SHOWSTATUS TRUE STATUSONTOP TRUE IN PROJECT "MicroStrategy Tutorial";

		ALTER PROJECT CONFIGURATION SHOWSTATUS TRUE IN PROJECT "MicroStrategy Tutorial";

		ALTER PROJECT CONFIGURATION STATUSONTOP FALSE IN PROJECT "MicroStrategy Tutorial";
		
#>
       
[CmdletBinding()]
param(
[string]$FileName,
[string]$ProjectDescription,
[string]$WarehouseName,
[string]$StatusHtmlInputFile,
[Switch]$ShowStatus,
[Switch]$StatusOnTop,
[string]$DocDirectory,
[int]$MaxNoAttrElems, 
[Switch]$UseWHLoginExec,
[Switch]$EnableObjectDeletion,
[int]$MaxReportExecTimeInSecs,
[int] $MaxScheduleReportExecTimeInSecs,
[int] $MaxNoIntelligentCubeResultRows,
[int] $MaxNoDatamartResultRows,
[int] $MaxNoReportResultRows,
[int] $MaxNoElemRows,
[int] $MaxNoIntResultRows,
[int] $MaxJobsUserAcct,
[int] $MaxJobsUserSession,
[int] $MaxExecJobsUser,
[int] $MaxJobsProject,
[int] $MaxUserSessionsProject,
[int] $MaxInteractiveSessionsPerUser,
[int] $MaxHistlistSubscriptions,
[int] $MaxCacheUpdateSubscriptions,
[int] $MaxEmailSubscriptions, 
[int] $MaxFileSubscriptions, 
[int] $MaxPrintSubscriptions, 
[int] $MaxMobileSubscriptions,
[int] $MaxFileSizeImportInMB,
[int] $MaxQuotaImportInMB,
[string] $ProjDrillMap,
[string] $DrillMapPath,
[string] $ReportTemplate,
[Switch] $ReportShowEmptyTemplate,
[string] $TemplateTemplate,
[Switch] $TemplateShowemptyTemplate,
[string] $MetricTemplate,
[Switch] $MetricShowEmptyTemplate,
[string] $DocumentTemplate, 
[Switch] $DocShowEmptytTemplate,  
[Switch] $UseDocWizard, 
[string] $IntelligentCubeFilePath,
[int] $IntelligentCubeMaxRamInMB,
[int] $IntelligentCubeMaxNumCubes,
[Switch] $DrillImmediate,
[Switch] $SortDrillPaths,
[string] $PdfHeaderLeft,
[string] $PdfHeaderCenter,
[string] $PdfHeaderRight,
[string] $PdfFooterLeft, 
[string] $PdfFooterCenter,
[string] $PdfFooterRight, 
[string] $ExportHeader, 
[string] $ExportFooter, 
[Switch] $ExportCompatibilityMode,
[Switch] $MergeHeaders,
[ValidateSet("PDF","MHT")]
[string] $ExportFlashFormat,
[Switch] $CreateUserProfileAtLogin,
[Switch] $ReportDescription,
[Switch] $PromptDetails,
[Switch] $FilterDetails, 
[Switch] $TemplateDetails, 
[Switch] $IncludePromptTitle,
[ValidateSet("Default", "Blank", "PromptNotAnswered", "NoSelection", "AllNone")] 
[string] $ReplaceUnansweredPrompt, 
[ValidateSet("No", "Yes", "Repeated")] 
[string] $ShowAttributeNameInPrompt,
[Switch] $IncludeUnUsedPrompts,
[ValidateSet("No", "Yes", "Automatic")] 
[string] $UseDelimitersInReportObjects,
[Switch] $UseAliasesInFiltersDetails,
[Switch] $ReportFilter, 
[ValidateSet("No", "Yes", "Automatic")] 
[string] $ReportFilterName,
[Switch] $ReportFilterDesc,
[Switch] $ReportLimits,
[Switch] $ViewFilter, 
[Switch] $MetricQualificationViewFilter,
[Switch] $DrillFilter,
[Switch] $SecurityFilter,
[Switch] $IncludeFilterType,
[Switch] $ShowEmptyExpression,
[Switch] $NewLineAfterTypeName,
[Switch] $NewLineBetweenFilter,
[ValidateSet("BeforeViewFilter", "AfterViewFilter")] 
[string] $ShowReportLimits,
[ValidateSet("ShowDefinition", "ShowName", "ShowNameAndDefinition")] 
[string] $ExpandShorcutFilters, 
[ValidateSet("No", "Yes", "Repeated")] 
[string]$ShowAttributeInListConditions,
[string] $SeparatorAfterAttrName,
[Switch] $NewLineAfterAttrName,
[ValidateSet("Or", "And", "Comma", "Custom")] 
[string] $SeparatorBetweenLastTwoElements,
[string] $CustomSeparatorBetweenLastTwoElements,
[Switch] $NewLineBetweenElements,
[Switch] $TrimElements, 
[ValidateSet("Symbols", "Names")] 
[string] $DisplayOperators, 
[Switch] $IncludeAttrFormInCondition,
[ValidateSet("Default", "ShowDate", "ShowExpressions")]  
[string] $DynamicDate,
[ValidateSet("No", "Yes", "Automatic")] 
[string] $NewLineBetweenConditions, 
[ValidateSet("Singlespace", "Doublespace")]
[string] $LineSpacing, 
[ValidateSet("No", "Yes", "Automatic")]  
[string] $ParenthesesAroundConditions, 
[ValidateSet("No", "Yes", "And", "Or")] 
[string] $LogicalOpBetweenConditions,
[ValidateSet("View", "Base")] 
[string] $UnitsFrom,
[ValidateSet("No", "Yes", "Automatic")]
[string] $BaseTemplate,
[Switch] $TemplateDescription,
[Switch] $NonMetricTemplateUnits,
[Switch] $Metrics,
[Switch] $OnlyConditional,
[Switch] $Formula,
[Switch] $Dimensionality,
[Switch] $Conditionality,
[Switch] $Transformation,
[ValidateSet("NoWatermark", "TextWatermark", "ImageWatermark")]
[string] $Watermark,
[string] $TextwatermarkText,
[Switch] $TextwatermarkSizeAutomatically,
[Switch] $TextwatermarkWashout,
[ValidateSet("Diagonal", "Horizontal")]
[string] $TextwatermarkOrientation,
[string] $ImagewatermarkSource,
[int] $ImagewatermarkScalePercentage, 
[Switch] $ImagewatermarkScalePercentageAuto,
[Switch] $AllowDocumentOverwriteWatermark, 
[string] $WebServerMacroWebserverAddress,  
[Switch] $EnableLinksInMht,
[ValidateSet("ValuePrompt", "ObjectPrompt", "AttributePrompt", "HierarchyPrompt", "AttributequalPrompt", "MetricQualification")]
[string[]] $CustomPromptStyles,
[string[]] $CustomPromptName, 
[string[]] $CustomPromptDescription,
[ValidateSet("ValuePrompt", "ObjectPrompt", "AttributePrompt", "HierarchyPrompt", "AttributequalPrompt", "MetricQualification")]
[string[]] $CustomPromptBaseStyle, 
[Switch] $DisplayAttrInAlphabeticalOrder,
[Switch] $EnablePersonalAnswers,
[Switch] $ExportPdfInHebrew,
[string] $DefaultDatamartDbInstanceName, 
[int] $WaitTimeForPromptAnswers, 
[int] $MaxWarehouseExecutionTime,
[int] $MaxSqlGenerationMemoryInMb, 
[int] $MaxNumInteractiveJobsPerProject, 
[ValidateSet("None", "Low", "High")] 
[string] $CacheEncryption,
[Switch] $DisableAutomaticExpirationDynamicDates,
[Switch] $ReRunMobileSubscriptions,
[Switch] $ReRunEmailSubscriptions,
[Switch] $DisableCreatingMatchingCaches,
[Switch] $KeepDocumentForManipulation,
[Switch] $CreateCubesByDatabaseConnection,
[Switch] $LoadCubesOnStartup, 
[Switch] $LoadCubesOnPublication,
[ValidateSet("Union", "Intersect")] 
[string] $SecurityFilterMerger, 
[Switch] $AttrShowNumeric,
[Switch] $AttrShowCharacter,
[Switch] $AttrShowBinary, 
[Switch] $AttrShowDate, 
[Switch] $AttrShowBigDecimal, 
[Switch] $AttrReplaceUnderscore,
[Switch] $AttrRemoveId,
[Switch] $AttrFirstLetterUppercase,
[string] $AttrIdStringPrefix, 
[string] $AttrDescStringPrefix, 
[string] $AttrLookupStringPrefix, 
[Switch] $FactShowNumeric,
[Switch] $FactShowCharacter, 
[Switch] $FactShowBinary, 
[Switch] $FactShowDate, 
[Switch] $FactShowbigDecimal,
[Switch] $FactReplaceUnderscore,
[Switch] $FactFirstLetterUppercase,
[Switch] $VerifyCustomColumnName, 
[string] $NullDisplayWarehouseStringToDisplay,
[string] $NullDisplayCrossTabulatorStringToDisplay, 
[Switch] $ReplaceNullWhenSorted,
[string] $ReplaceNullWhenSortedValue,
[string] $NotCalculatedDisplayStringToDisplay,
[string] $MissingDisplayStringToDisplay, 
[Switch] $OverrideGraphTemplateFont,
[string] $Characterset,
[string] $FontName,
[Switch] $UseZeroForGraphNull,
[ValidateSet("None", "Standard", "Optimize")] 
[string] $GraphRoundedEffect,
[string] $EmptyReportMessage,
[Switch] $DisplayEmptyMessageInDocument,
[Switch] $KeepPageByWhenSaving,
[ValidateSet("Allow", "Allowwithwarning", "Prevent")] 
[string] $OverwriteWithOlapReports,
[Switch] $MoveSortKeysWithPivot,
[Switch] $AppendEmailFooter,
[string] $EmailFooterText, 
[Switch] $EnablePrintingRange,
[Switch] $DeliverSubscriptionWithNoData,
[Switch] $DeliverSubscriptionWithPartialResults,
[Switch] $AllSchedules,
[string[]] $ScheduleName,  
[Parameter(Mandatory=$True)]
[string] $ProjectName
)

$CmdLine = $null
$Empty = [String]::Empty
$CmdLine = "ALTER PROJECT CONFIGURATION "
switch ($true) { 
  ($ProjectDescription -ne $Empty)     		{ $CmdLine += "DESCRIPTION `"$ProjectDescription`" " }
  ($WarehouseName -ne $Empty)				{ $CmdLine += "WAREHOUSE `"$WarehouseName`" " }
  ($StatusHtmlInputFile -ne $Empty)		   	{ $CmdLine += "STATUS `"$StatusHtmlInputFile`" " }
  $ShowStatus								{ $CmdLine += "SHOWSTATUS TRUE " }
  $StatusOnTop			 					{ $CmdLine += "STATUSONTOP TRUE " }
  ($DocDirectory  -ne $Empty)				{ $CmdLine += "DOCDIRECTORY `"$DocDirectory`" " }
  ($MaxNoAttrElems -ne 0)				{ $CmdLine += "MAXNOATTRELEMS $MaxNoAttrElems " }
  $UseWHLoginExec							{ $CmdLine += "USEWHLOGINEXEC TRUE " }
  $EnableObjectDeletion						{ $CmdLine += "ENABLEOBJECTDELETION TRUE " }
  ($MaxReportExecTimeInSecs -ne 0)		{ $CmdLine += "MAXREPORTEXECTIME $MaxReportExecTimeInSecs "} 
  ($MaxScheduleReportExecTimeInSecs -ne 0) { $CmdLine += "MAXSCHEDULEREPORTEXECTIME $MaxScheduleReportExecTimeInSecs "}
  ($MaxNoIntelligentCubeResultRows -ne 0)  { $CmdLine += "MAXNOINTELLIGENTCUBERESULTROWS $MaxNoIntelligentCubeResultRows " }
  ($MaxNoDatamartResultRows -ne 0)         { $CmdLine += "MAXNODATAMARTRESULTROWS $MaxNoDatamartResultRows " }
  ($MaxNoReportResultRows -ne 0)  		{ $CmdLine += "MAXNOREPORTRESULTROWS $MaxNoReportResultRows " }
  ($MaxNoElemRows -ne 0) 				{ $CmdLine += "MAXNOELEMROWS $MaxNoElemRows " }
  ($MaxNoIntResultRows -ne 0)  			{ $CmdLine += "MAXNOINTRESULTROWS $MaxNoIntResultRows "}
  ($MaxJobsUserAcct -ne 0)              { $CmdLine += "MAXJOBSUSERACCT $MaxJobsUserAcct "}
  ($MaxJobsUserSession -ne 0)           { $CmdLine += "MAXJOBSUSERSESSION $MaxJobsUserSession " }
  ($MaxExecJobsUser -ne 0)              { $CmdLine += "MAXEXECJOBSUSER $MaxExecJobsUser "}
  ($MaxJobsProject -ne 0)               { $CmdLine += "MAXJOBSPROJECT $MaxJobsProject "}
  ($MaxUserSessionsProject -ne 0)       { $CmdLine += "MAXUSERSESSIONSPROJECT $MaxUserSessionsProject "}
  ($MaxInteractiveSessionsPerUser -ne 0)   { $CmdLine += "MAXINTERACTIVESESSIONSPERUSER $MaxInteractiveSessionsPerUser " }
  ($MaxHistlistSubscriptions -ne 0)        { $CmdLine += "MAXHISTLISTSUBSCRIPTIONS $MaxHistlistSubscriptions " }
  ($MaxCacheUpdateSubscriptions -ne 0)     { $CmdLine += "MAXCACHEUPDATESUBSCRIPTIONS $MaxCacheUpdateSubscriptions " }
  ($MaxEmailSubscriptions -ne 0)     	   { $CmdLine += "MAXEMAILSUBSCRIPTIONS $MaxEmailSubscriptions " }
  ($MaxFileSubscriptions -ne 0)            { $CmdLine += "MAXFILESUBSCRIPTIONS $MaxFileSubscriptions " }
  ($MaxPrintSubscriptions -ne 0)     { $CmdLine += "MAXPRINTSUBSCRIPTIONS $MaxPrintSubscriptions " }
  ($MaxMobileSubscriptions -ne 0)    { $CmdLine += "MAXMOBILESUBSCRIPTIONS $MaxMobileSubscriptions " }
  ($MaxFileSizeImportInMB -ne 0)     { $CmdLine += "MAXFILESIZEIMPORT $MaxFileSizeImportInMB " }
  ($MaxQuotaImportInMB -ne 0)        { $CmdLine += "MAXQUOTAIMPORT $MaxQuotaImportInMB " }
  ($ProjDrillMap -ne $Empty)             { $CmdLine += "PROJDRILLMAP `"$ProjDrillMap`" " }
  ($DrillMapPath -ne $Empty)             { $CmdLine += "IN FOLDER `"$DrillMapPath`" " }
  ($ReportTemplate -ne $Empty)           { $CmdLine += "REPORTTPL `"$ReportTemplate`" " }
  $ReportShowEmptyTemplate               { $CmdLine += "REPORTSHOWEMPTYTPL TRUE " }
  ($TemplateTemplate -ne $Empty)         { $CmdLine += "TEMPLATETPL `"$TemplateTemplate`" " }
  $TemplateShowemptyTemplate             { $CmdLine += "TEMPLATESHOWEMPTYTPL TRUE " }
  ($MetricTemplate -ne $Empty)           { $CmdLine += "METRICTPL `"$MetricTemplate`" " }
  $MetricShowEmptyTemplate               { $CmdLine += "METRICSHOWEMPTYTPL TRUE " }
  ($DocumentTemplate -ne $Empty)         { $CmdLine += "DOCTPL `"$DocumentTemplate`" " }
  $DocShowEmptytTemplate                 { $CmdLine += "DOCSHOWEMPTYTPL TRUE " }
  $UseDocWizard					         { $CmdLine += "USEDOCWIZARD TRUE " }
  ($IntelligentCubeFilePath -ne $Empty)  { $CmdLine += "INTELLIGENTCUBEFILEDIR `"$IntelligentCubeFilePath`" " }
  ($IntelligentCubeMaxRamInMB -ne 0) { $CmdLine += "INTELLIGENTCUBEMAXRAM $IntelligentCubeMaxRamInMB " }
  ($IntelligentCubeMaxNumCubes -ne 0) { $CmdLine += "INTELLIGENTCUBEMAXNO $IntelligentCubeMaxNumCubes " }
  $DrillImmediate                         { $CmdLine += "DRILLIMMEDIATE TRUE " }
  $SortDrillPaths						  { $CmdLine += "SORTDRILLPATHS TRUE " }
  ($PdfHeaderLeft -ne $Empty)             { $CmdLine += "PDFHEADERLEFT `"$PdfHeaderLeft`" " }
  ($PdfHeaderCenter -ne $Empty)           { $CmdLine += "PDFHEADERCENTER `"$PdfHeaderCenter`" " }
  ($PdfHeaderRight -ne $Empty)            { $CmdLine += "PDFHEADERRIGHT `"$PdfHeaderRight`" " }
  ($PdfFooterCenter -ne $Empty)           { $CmdLine += "PDFFOOTERCENTER `"$PdfFooterCenter`" " }
  ($PdfFooterRight -ne $Empty)            { $CmdLine += "PDFFOOTERRIGHT `"$PdfFooterRight`" " }
  ($ExportHeader -ne $Empty)              { $CmdLine += "EXPORTHEADER `"$ExportHeader`" " }
  ($ExportFooter -ne $Empty)              { $CmdLine += "EXPORTFOOTER `"$ExportFooter`" " }
  $ExportCompatibilityMode                { $CmdLine += "EXPORTCOMPATIBILITYMODE TRUE " }
  $MergeHeaders				              { $CmdLine += "MERGEHEADERS TRUE " }
  ($ExportFlashFormat -ne $Empty)         { $CmdLine += "EXPORTFLASHFORMAT `"$ExportFlashFormat`" " }
  $CreateUserProfileAtLogin				  { $CmdLine += "CREATEUSERPROFILEATLOGIN TRUE " }
  $ReportDescription					  { $CmdLine += "REPORTDESCRIPTION TRUE " }
  $PromptDetails						  { $CmdLine += "PROMPTDETAILS TRUE " }
  $FilterDetails						  { $CmdLine += "FILTERDETAILS TRUE " }
  $TemplateDetails						  { $CmdLine += "TEMPLATEDETAILS TRUE " }
  $IncludePromptTitle					  { $CmdLine += "INCLUDEPROMPTTITLE TRUE " }
  ($ReplaceUnansweredPrompt -ne $Empty)   { $CmdLine += "REPLACEUNANSWEREDPROMPT `"$ReplaceUnansweredPrompt`" " }
  ($ShowAttributeNameInPrompt -ne $Empty) { $CmdLine += "SHOWATTRIBUTENAMEINPROMPT `"$ShowAttributeNameInPrompt`" " }
  $IncludeUnUsedPrompts                   { $CmdLine += "INCLUDEUNUSEPROMPTS TRUE " }
  ($UseDelimitersInReportObjects -ne $Empty)     { $CmdLine += "USEDELIMITERSINREPORTOBJECTS $UseDelimitersInReportObjects " }
  $UseAliasesInFiltersDetails                    { $CmdLine += "USEALIASESINFILTERSDETAILS TRUE " }
  $ReportFilter								{ $CmdLine += "REPORTFILTER TRUE " } 
  ($ReportFilterName -ne $Empty)			{ $CmdLine += "REPORTFILTERNAME $ReportFilterName " }  
$ReportFilterDesc  							{ $CmdLine += "REPORTFILTERDESC TRUE " }
$ReportLimits 								{ $CmdLine += "REPORTLIMITS TRUE " }
$ViewFilter 								{ $CmdLine += "VIEWFILTER TRUE " }
$MetricQualificationViewFilter				{ $CmdLine += "METRICQUALIFICATIONVIEWFILTER TRUE " }
$DrillFilter								{ $CmdLine += "DRILLFILTER TRUE " }
$SecurityFilter 							{ $CmdLine += "SECURITYFILTER TRUE " }
$IncludeFilterType							{ $CmdLine += "INCLUDEFILTERTYPE TRUE " }
$ShowEmptyExpression						{ $CmdLine += "SHOWEMPTYEXPRESSION TRUE " }
$NewLineAfterTypeName						{ $CmdLine += "NEWLINEAFTERTYPENAME TRUE " }
$NewLineBetweenFilter 						{ $CmdLine += "NEWLINEBETWEENFILTER TRUE " }
($ShowReportLimits -ne $Empty)				{ $CmdLine += "SHOWREPORTLIMITS $ShowReportLimits "}
($ExpandShorcutFilters -ne $Empty)          { $CmdLine += "EXPANDSHORCUTFILTERS $ExpandShorcutFilters "}
($ShowAttributeInListConditions -ne $Empty) { $CmdLine += "SHOWATTRIBUTEINLISTCONDITIONS $ShowAttributeInListConditions "}
($SeparatorAfterAttrName -ne $Empty)        { $CmdLine += "SEPARATORAFTERATTRNAME `"$SeparatorAfterAttrName`" "} 
$NewLineAfterAttrName						{ $CmdLine += "NEWLINEAFTERATTRNAME TRUE "} 
($SeparatorBetweenLastTwoElements -ne $Empty)        { $CmdLine += "SEPARATORBETWEENLASTTWOELEMENTS $SeparatorBetweenLastTwoElements "}
($CustomSeparatorBetweenLastTwoElements -ne $Empty)  { $CmdLine += "`"$CustomSeparatorBetweenLastTwoElements`" " }
$NewLineBetweenElements						{ $CmdLine += "NEWLINEBETWEENELEMENTS TRUE "}
$TrimElements								{ $CmdLine += "TRIMELEMENTS TRUE "} 
($DisplayOperators -ne $Empty)				{ $CmdLine += "DISPLAYOPERATORS $DisplayOperators "}
$IncludeAttrFormInCondition					{ $CmdLine += "INCLUDEATTRFORMINCONDITION TRUE "} 
($DynamicDate -ne $Empty) 					{ $CmdLine += "DYNAMICDATE $DynamicDate "}
($NewLineBetweenConditions -ne $Empty) 		{ $CmdLine += "NEWLINEBETWEENCONDITIONS $NewLineBetweenConditions "}
($LineSpacing -ne $Empty) 					{ $CmdLine += "LINESPACING $LineSpacing "}
($ParenthesesAroundConditions -ne $Empty)	{ $CmdLine += "PARENTHESESAROUNDCONDITIONS $ParenthesesAroundConditions "}
($LogicalOpBetweenConditions -ne $Empty)	{ $CmdLine += "LOGICALOPBETWEENCONDITIONS $LogicalOpBetweenConditions "}
($UnitsFrom -ne $Empty)						{ $CmdLine += "UNITSFROM $UnitsFrom "}
($BaseTemplate -ne $Empty)					{ $CmdLine += "BASETEMPLATE $BaseTemplate "}
$TemplateDescription						{ $CmdLine += "TEMPLATEDESCRIPTION TRUE "}
$NonMetricTemplateUnits						{ $CmdLine += "NONMETRICTEMPLATEUNITS TRUE "}
$Metrics									{ $CmdLine += "METRICS TRUE "}
$OnlyConditional							{ $CmdLine += "ONLYCONDITIONAL TRUE "} 
$Formula									{ $CmdLine += "FORMULA TRUE "} 
$Dimensionality								{ $CmdLine += "DIMENSIONALITY TRUE "} 
$Conditionality								{ $CmdLine += "CONDITIONALITY TRUE "} 
$Transformation								{ $CmdLine += "TRANSFORMATION TRUE "}
($Watermark -ne $Empty -and $Watermark -ne $true) 	{ $CmdLine += "NOWATERMARK "}
($TextwatermarkText -ne $Empty )            		{ $CmdLine += "TEXTWATERMARK TEXT `"$TextwatermarkText`" "}
$TextwatermarkSizeAutomatically						{ $CmdLine += "SIZEAUTOMATICALLY TRUE " } 
$TextwatermarkWashout								{ $CmdLine += "WASHOUT TRUE "}
($TextwatermarkOrientation -ne $Empty)				{ $CmdLine += "ORIENTATION $TextwatermarkOrientation "}
($ImagewatermarkSource -ne $Empty)					{ $CmdLine += "IMAGEWATERMARK SOURCE `"$ImagewatermarkSource`" "}
($ImagewatermarkScalePercentage -ne 0)				{ $CmdLine += "SCALE $ImagewatermarkScalePercentage "}
$ImagewatermarkScalePercentageAuto					{ $CmdLine += "AUTO "}
$AllowDocumentOverwriteWatermark            		{ $CmdLine += "ALLOWDOCUMENTOVERWRITEWATERMARK TRUE "}  
($WebServerMacroWebserverAddress -ne $Empty)    	{ $CmdLine += "WEBSERVERMACRO `"$WebServerMacroWebserverAddress`" "} 
$EnableLinksInMht									{ $CmdLine += "ENABLELINKSINMHT TRUE "}  
($CustomPromptStyles -ne $null)             { 
	$c = 0
	$CmdLine += "CUSTOMPROMPSTYLES "
	foreach($Style in $CustomPromptStyles) {
	if($c -gt 0) { $CmdLine += ", " }
    $CmdLine += "$Style " 
    if($CustomPromptName[$c] -ne $null)		         { $CmdLine += "NAME `"$($CustomPromptName[$c])`" "}
    if($CustomPromptDescription[$c] -ne $null)       { $CmdLine += "DESCRIPTION `"$($CustomPromptDescription[$c])`" "}
    if($CustomPromptBaseStyle[$c] -ne $null)		 { $CmdLine += "BASESTYLE `"$($CustomPromptBaseStyle[$c])`" "; }
	$c++
	}
}
$DisplayAttrInAlphabeticalOrder              { $CmdLine += "DISPLAYATTRINALPHABETICALORDER TRUE "} 
$EnablePersonalAnswers 						 { $CmdLine += "ENABLEPERSONALANSWERS TRUE " } 
$ExportPdfInHebrew							 { $CmdLine += "EXPORTPDFINHEBREW TRUE " }
($DefaultDatamartDbInstanceName -ne $Empty)  { $CmdLine += "DEFAULTDATAMART `"$DefaultDatamartDbInstanceName`" " } 
($WaitTimeForPromptAnswers -ne 0)		 { $CmdLine += "WAITTIMEFORPROMPTANSWERS $WaitTimeForPromptAnswers " } 
($MaxWarehouseExecutionTime -ne 0)  	 { $CmdLine += "WAREHOUSEEXECUTIONTIME $MaxWarehouseExecutionTime " }
($MaxSqlGenerationMemoryInMb -ne 0) 	 { $CmdLine += "SQLGENERATIONMEMORY $MaxSqlGenerationMemoryInMb " }
($MaxNumInteractiveJobsPerProject -ne 0) { $CmdLine += "INTERACTIVEJOBPERPROJECT $MaxNumInteractiveJobsPerProject "}
($CacheEncryption -ne $Empty)				 { $CmdLine += "CACHEENCRYPTION $CacheEncryption " }
$DisableAutomaticExpirationDynamicDates		 { $CmdLine += "DISABLEAUTOMATICEXPIRATIONDYNAMICDATES TRUE " }
$ReRunMobileSubscriptions					 { $CmdLine += "RERUNMOBILESUBSCRIPTIONS TRUE " }
$ReRunEmailSubscriptions					 { $CmdLine += "RERUNEMAILSUBSCRIPTIONS TRUE " }
$DisableCreatingMatchingCaches				 { $CmdLine += "DISABLECREATINGMATCHINGCACHES TRUE " }
$KeepDocumentForManipulation				 { $CmdLine += "KEEPDOCUMENTFORMANIPULATION TRUE " }
$CreateCubesByDatabaseConnection			 { $CmdLine += "CREATECUBESBYDATABASECONNECTION TRUE " }
$LoadCubesOnStartup							 { $CmdLine += "LOADCUBESONSTARTUP TRUE " }
$LoadCubesOnPublication						 { $CmdLine += "LOADCUBESONPUBLICATION TRUE " }
($SecurityFilterMerger -ne $Empty)           { $CmdLine += "SECURITYFILTERMERGER $SecurityFilterMerger " } 
$AttrShowNumeric							 { $CmdLine += "ATTRSHOWNUMERIC TRUE " }
$AttrShowCharacter							 { $CmdLine += "ATTRSHOWCHARACTER TRUE " }
$AttrShowBinary								 { $CmdLine += "ATTRSHOWBINARY TRUE " }
$AttrShowDate								 { $CmdLine += "ATTRSHOWDATE TRUE " }
$AttrShowBigDecimal							 { $CmdLine += "ATTRSHOWBIGDECIMAL TRUE " }
$AttrReplaceUnderscore						 { $CmdLine += "ATTRREPLACEUNDERSCORE TRUE " }
$AttrRemoveId								 { $CmdLine += "ATTRREMOVEID TRUE " }
$AttrFirstLetterUppercase					 { $CmdLine += "ATTRFIRSTLETTERUPPERCASE TRUE " }
($AttrIdStringPrefix -ne $Empty)			 { $CmdLine += "ATTRIDSTRING `"$AttrIdStringPrefix`" "} 
($AttrDescStringPrefix -ne $Empty)           { $CmdLine += "ATTRDESCSTRING `"$AttrDescStringPrefix`" " } 
($AttrLookupStringPrefix -ne $Empty)		 { $CmdLine += "ATTRLOOKUPSTRING `"$AttrLookupStringPrefix`" "}
$FactShowNumeric						     { $CmdLine += "FACTSHOWNUMERIC TRUE " } 
$FactShowCharacter							 { $CmdLine += "FACTSHOWCHARACTER TRUE " }
$FactShowBinary								 { $CmdLine += "FACTSHOWBINARY TRUE " }
$FactShowDate								 { $CmdLine += "FACTSHOWDATE TRUE " }
$FactShowbigDecimal							 { $CmdLine += "FACTSHOWBIGDECIMAL TRUE " }
$FactReplaceUnderscore						 { $CmdLine += "FACTREPLACEUNDERSCORE TRUE " }
$FactFirstLetterUppercase					 { $CmdLine += "FACTFIRSTLETTERUPPERCASE TRUE " }
$VerifyCustomColumnName						 { $CmdLine += "VERIFYCUSTOMCOLUMNNAME TRUE " }
($NullDisplayWarehouseStringToDisplay -ne $Empty)       { $CmdLine += "NULLDISPLAYWAREHOUSE `"$NullDisplayWarehouseStringToDisplay`" " } 
($NullDisplayCrossTabulatorStringToDisplay -ne $Empty)  { $CmdLine += "NULLDISPLAYCROSSTABULATOR `"$NullDisplayCrossTabulatorStringToDisplay`" " }
$ReplaceNullWhenSorted									{ $CmdLine += "REPLACENULLWHENSORTED TRUE " }
($ReplaceNullWhenSortedValue -ne $Empty)				{ $CmdLine += "REPLACENULLWHENSORTEDVALUE  $ReplaceNullWhenSortedValue " }
($NotCalculatedDisplayStringToDisplay -ne $Empty)       { $CmdLine += "NOTCALCULATEDDISPLAY `"$NotCalculatedDisplayStringToDisplay`" " }
($MissingDisplayStringToDisplay -ne $Empty)			    { $CmdLine += "MISSINGDISPLAY `"$MissingDisplayStringToDisplay`" " }
$OverrideGraphTemplateFont								{ $CmdLine += "OVERRIDEGRAPHTEMPLATEFONT TRUE " }
($Characterset -ne $Empty)								{ $CmdLine += "CHARACTERSET `"$Characterset`" " }
($FontName -ne $Empty)									{ $CmdLine += "FONT `"$FontName`" " }
$UseZeroForGraphNull									{ $CmdLine += "USEZEROFORGRAPHNULL TRUE " }
($GraphRoundedEffect -ne $Empty)						{ $CmdLine += "GRAPHROUNDEDEFFECT $GraphRoundedEffect " }
($EmptyReportMessage -ne $Empty)						{ $CmdLine += "EMPTYREPORTMESSAGE `"$EmptyReportMessage`" " }
$DisplayEmptyMessageInDocument							{ $CmdLine += "DISPLAYEMPTYMESSAGEINDOCUMENT TRUE " }
$KeepPageByWhenSaving									{ $CmdLine += "KEEPPAGEBYWHENSAVING TRUE " }
($OverwriteWithOlapReports -ne $Empty)					{ $CmdLine += "OVERWRITEWITHOLAPREPORTS $OverwriteWithOlapReports "}
$MoveSortKeysWithPivot									{ $CmdLine += "MOVESORTKEYSWITHPIVOT TRUE "}
$AppendEmailFooter										{ $CmdLine += "APPENDEMAILFOOTER TRUE " }
($EmailFooterText -ne $Empty)							{ $CmdLine += "EMAILFOOTERTEXT `"$EmailFooterText`" " }
$EnablePrintingRange									{ $CmdLine += "ENABLEPRINTINGRANGE TRUE "}
$DeliverSubscriptionWithNoData							{ $CmdLine += "DELIVERSUBSCRIPTIONWITHNODATA TRUE "}
$DeliverSubscriptionWithPartialResults					{ $CmdLine += "DELIVERSUBSCRIPTIONWITHPARTIALRESULTS TRUE " }
$AllSchedules											{ $CmdLine += "ALLSCHEDULES " } 
($ScheduleName -ne $null)								{ 
	$CmdLine += "SCHEDULES "
	$c = 0
	foreach($schedule in $ScheduleName) { $CmdLine += "`"$schedule`", " }
	$CmdLine = $CmdLine.TrimEnd(", ")
}
($ProjectName -ne $Empty)			   		           { $CmdLine += "IN PROJECT `"$ProjectName`";"}
 }

if($FileName -ne $Empty) { Out-File -FilePath $FileName -Encoding ascii -InputObject $CmdLine } else { $CmdLine }
}