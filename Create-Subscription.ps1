
Function Create-Subscription {
<#
		.NAME
		 Create-Subscription
		.DESCRIPTION
         Creates a MicroStrategy Subscription.
        .NOTES
		 When creating a subscription with the recipient being only an ADDRESS it will be an address of the user who is creating the subscription (i.e. the user connected to MicroStrategy Command Manager)
		
		 Subscriptions for prompted reports and documents are supported in a limited basis. Embedded prompts or prompts within prompts are not supported in this release. If the answer is not provided for a prompt the default answer will be used. Otherwise the execution will fail.
					
         AUTHOR....:  Craig Grady
         LAST EDIT.:  01/12/2015
         CREATED...:  01/09/2015
		
		.Syntax
		MicroStrategy Command Manager "CREATE HISTORYLISTSUBSCRIPTION" syntax
		CREATE HISTORYLISTSUBSCRIPTION ["<subscription_name>"] [FOR OWNER "<login_name>"] SCHEDULE "<schedule_name>"  (USER "<login_name>" | GROUP "<user_group_name>")  CONTENT ("<report_or_document_name>" IN FOLDER "<location_name>" | GUID <report_or_document_guid>) [CONTENTTYPE (REPORT|DOCUMENT)] IN PROJECT "<project_name>"  [OVERWRITEOLDERVERSION (TRUE|FALSE)] [HISTORYLINKURL "<historylinkurl>"] [HISTORYLINKWEBSERVER "<historylinkwebserver>"] [EXPIRATIONDATE mm/dd/yyyy] [SENDNOTIFICATION (TRUE | FALSE)] [(SENDNOTIFICATION_ADDRESS "<address_name>" |PHYSICALADDRESS "<physical_address>" DEVICE "<device_name>")] [RUNFRESH (TRUE|FALSE)] [CREATEUPDATECACHE (TRUE|FALSE)] [MODIFICATIONBYRECIPIENTS (TRUE|FALSE)] [PROMPT "<prompt_name>" ANSWER "<prompt_answer>" [, PROMPT "<prompt_name>" ANSWER "<prompt_answer>" ...]]; 
		
		1st Example of MicroStrategy Command Manager "CREATE HISTORYLISTSUBSCRIPTION" syntax implemented
		CREATE HISTORYLISTSUBSCRIPTION "PromptedSubscription1" FOR OWNER "Administrator" SCHEDULE "Books Closed" USER "Administrator" CONTENT "Product Analysis Report Builder" IN FOLDER "Public Objects\Reports\MicroStrategy Platform Capabilities\Ad hoc Reporting\Prompts" IN PROJECT "MicroStrategy Tutorial" PROMPT "Select attributes Time-Product-Geography" ANSWER "8D679D5111D3E4981000E787EC6DE8A4,8D679D3711D3E4981000E787EC6DE8A4" , PROMPT "Sales Metrics" ANSWER "8D679D5111D3E4981000E787EC6DE8A4,4C051DB611D3E877C000B3B2D86C964F";

		2nd Example of MicroStrategy Command Manager "CREATE HISTORYLISTSUBSCRIPTION" syntax implemented
		#CREATE HISTORYLISTSUBSCRIPTION "My first history list subscription" FOR OWNER "demo" SCHEDULE "Books Closed" USER "demo" CONTENT "Customer Income Analysis" IN FOLDER "\Public Objects\Reports\Subject Areas\Customer Analysis"  IN PROJECT "Microstrategy Tutorial" OVERWRITEOLDERVERSION TRUE CREATEUPDATECACHE FALSE HISTORYLINKURL "samplehistorylinkurl" HISTORYLINKWEBSERVER "samplehistorywebserver";

		MicroStrategy Command Manager "CREATE EMAILSUBSCRIPTION" syntax
		CREATE EMAILSUBSCRIPTION ["<subscription_name>"] [FOR OWNER "<login_name>"] SCHEDULE "<schedule_name>" (ADDRESS "<address_name>" | USER "<login_name>" | CONTACT "<contact_name>" [ADDRESS "<address_name>"] | CONTACTGROUP "<contact_group_name>") CONTENT ("<report_or_document_name>" IN FOLDER "<location_name>" | GUID <report_or_document_guid>) [CONTENTTYPE (REPORT | DOCUMENT)] IN PROJECT "<project_name>" SUBJECT "<subject>" DELIVERYFORMAT (HTML | CSV FILENAME "<file_name>" | PLAINTEXT FILENAME "<file_name>" [TEXTDELIMITER (COMMA | SPACE | TAB | SEMICOLON | COLON | OTHER "<custom_delimiter>")] | EXCEL FILENAME "<file_name>" | PDF FILENAME "<file_name>" | FLASH FILENAME "<file_name>") [COMPRESSCONTENTS ZIPFILENAME "<zip_file_name>" [ZIPFILEPASSWORD "<zip_file_pasword>"]] [MESSAGE "<message>"]  [EXPIRATIONDATE mm/dd/yyyy] [SENDTOHISTORYLIST (TRUE | FALSE)] [INCLUDELINK (TRUE | FALSE)] [HISTORYLINKWEBSERVER "<historylinkwebserver>"] [SENDPREVIEWNOW (TRUE | FALSE)] [INCLUDEDATA (TRUE|FALSE)] [MODIFICATIONBYRECIPIENTS (TRUE | FALSE)] [PROMPT "<prompt_name>" ANSWER "<prompt_answer>" [,PROMPT "<prompt_name>" ANSWER "<prompt_answer>"...]];
		
		1st Example of MicroStrategy Command Manager "CREATE EMAILSUBSCRIPTION" syntax implemented
		CREATE EMAILSUBSCRIPTION "My first email subscription" FOR OWNER "demo" SCHEDULE "Books Closed"  CONTACT "Contact1" CONTENT "Customer Income Analysis" IN FOLDER "\Public Objects\Reports\Subject Areas\Customer Analysis"  IN PROJECT "Microstrategy Tutorial"   EXPIRATIONDATE 05/11/2011 SUBJECT "My First Subscription" MESSAGE "My First Message" INCLUDELINK TRUE SENDPREVIEWNOW TRUE SENDTOHISTORYLIST TRUE INCLUDEDATA TRUE  DELIVERYFORMAT PDF FILENAME "MySubscription" COMPRESSCONTENTS ZIPFILENAME "zip.txt";
		
		MicroStrategy Command Manager "CREATE FILESUBSCRIPTION" syntax
		CREATE FILESUBSCRIPTION  ["<subscription_name>"]  [FOR OWNER "<login_name>"] SCHEDULE "<schedule_name>"  (ADDRESS "<address_name>" | USER "<login_name>"  | CONTACT "<contact_name>"  [ADDRESS "<address_name>"] | CONTACTGROUP "<contact_group_name>") CONTENT ("<report_or_document_name>" IN FOLDER "<location_name>" | GUID <report_or_document_guid>) [CONTENTTYPE (REPORT|DOCUMENT)] IN PROJECT "<project_name>"  DELIVERYFORMAT (CSV | PLAINTEXT [TEXTDELIMITER (COMMA | SPACE | TAB | SEMICOLON | COLON | OTHER "<custom_delimiter>")] | EXCEL | PDF | HTML | FLASH) [COMPRESSCONTENTS ZIPFILENAME "<zip_file_name>" [ZIPFILEPASSWORD "<zip_file_pasword>"]] [EXPIRATIONDATE mm/dd/yyyy] FILENAME "<file_name>"  [SENDPREVIEWNOW(TRUE|FALSE)] [SENDNOTIFICATION (TRUE | FALSE)] [(SENDNOTIFICATION_ADDRESS "<address_name>" |PHYSICALADDRESS "<physical_address>" DEVICE "<device_name>")] [MODIFICATIONBYRECIPIENTS (TRUE|FALSE)] [PROMPT "<prompt_name>" ANSWER "<prompt_answer>" [, PROMPT "<prompt_name>" ANSWER "<prompt_answer>" ...]];
		
		1st Example of MicroStrategy Command Manager "CREATE FILESUBSCRIPTION" syntax implemented
		CREATE FILESUBSCRIPTION "My first file subscription" SCHEDULE "Books Closed" CONTACT "Contact2" CONTENT "Customer Income Analysis" IN FOLDER "\Public Objects\Reports\Subject Areas\Customer Analysis"  IN PROJECT "Microstrategy Tutorial" DELIVERYFORMAT PDF FILENAME "MySubscription" COMPRESSCONTENTS ZIPFILENAME "Zip.zip"  EXPIRATIONDATE 05/11/2007 SENDPREVIEWNOW TRUE SENDNOTIFICATION TRUE;		
		
		MicroStrategy Command Manager "CREATE PRINTSUBSCRIPTION" syntax
		CREATE PRINTSUBSCRIPTION ["<subscription_name>"] [FOR OWNER "<login_name>"] SCHEDULE "<schedule_name>"  (ADDRESS "<address_name>" | USER "<login_name>"  | CONTACT "<contact_name>"  [ADDRESS "<address_name>"] | CONTACTGROUP "<contact_group_name>") CONTENT ("<report_or_document_name>" IN FOLDER "<location_name>" | GUID <report_or_document_guid>) [CONTENTTYPE (REPORT|DOCUMENT)] IN PROJECT "<project_name>"  [PRINTRANGE (PAGES FROM <start_page> TO <end_page> | ALL)] [COPIES <number_of_copies>] [COLLATE (TRUE|FALSE)] [SENDPREVIEWNOW(TRUE|FALSE)] [SENDNOTIFICATION (TRUE | FALSE)] [(SENDNOTIFICATION_ADDRESS "<address_name>"|PHYSICALADDRESS "<physical_address>" DEVICE "<device_name>")] [MODIFICATIONBYRECIPIENTS (TRUE|FALSE)] [PROMPT "<prompt_name>" ANSWER "<prompt_answer>" [, PROMPT "<prompt_name>" ANSWER "<prompt_answer>" ...]];
		
		1st Example of MicroStrategy Command Manager "CREATE PRINTSUBSCRIPTION" syntax implemented
		CREATE PRINTSUBSCRIPTION SCHEDULE "First of Month" CONTACT "Contact2"  CONTENT "Customer Income Analysis" IN FOLDER "\Public Objects\Reports\Subject Areas\Customer Analysis" IN PROJECT "MicroStrategy Tutorial"  PRINTRANGE PAGES FROM 1 TO 100 COPIES 10 COLLATE TRUE SENDPREVIEWNOW TRUE SENDNOTIFICATION TRUE;
		
		MicroStrategy Command Manager "CREATE CACHEUPDATESUBSCRIPTION" syntax
		CREATE CACHEUPDATESUBSCRIPTION ["<subscription_name>"] [FOR OWNER "<login_name>"] SCHEDULE "<schedule_name>"  (USER "<login_name>" | GROUP "<user_group_name>")  CONTENT ("<report_or_document_name>" IN FOLDER "<location_name>" | GUID <report_or_document_guid>) [CONTENTTYPE (REPORT|DOCUMENT)] IN PROJECT "<project_name>"  [DELIVERYFORMAT (EXCEL | PDF | HTML | FLASH)] [SENDNOTIFICATION (TRUE | FALSE)] [(SENDNOTIFICATION_ADDRESS "<address_name>" |PHYSICALADDRESS "<physical_address>" DEVICE "<device_name>")] [PROMPT "<prompt_name>" ANSWER "<prompt_answer>" [, PROMPT "<prompt_name>" ANSWER "<prompt_answer>" ...]];

		1st Example of MicroStrategy Command Manager "CREATE CACHEUPDATESUBSCRIPTION" syntax implemented
		CREATE CACHEUPDATESUBSCRIPTION "My first cache subscription" SCHEDULE "Monday Morning" USER "demo" CONTENT "Yearly Revenue Growth by Customer Region" IN FOLDER "\Public Objects\Reports\Subject Areas\Customer Analysis" IN PROJECT "MicroStrategy Tutorial" DELIVERYFORMAT HTML SENDNOTIFICATION FALSE;

		MicroStrategy Command Manager "CREATE MOBILESUBSCRIPTION" syntax
		CREATE MOBILESUBSCRIPTION ["<subscription_name>"] [FOR OWNER "<login_name>"] SCHEDULE "<schedule_name>"  (USER "<login_name>" | GROUP "<user_group_name>")  CONTENT ("<report_or_document_name>" IN FOLDER "<location_name>" | GUID <report_or_document_guid>) [CONTENTTYPE (REPORT|DOCUMENT)] IN PROJECT "<project_name>" [EXPIRATIONDATE mm/dd/yyyy] [RUNFRESH (TRUE|FALSE)] [CREATEUPDATECACHE (TRUE|FALSE)] [MODIFICATIONBYRECIPIENTS (TRUE|FALSE)] [DEVICE (PHONE | BLACKBERRY | TABLET)] [PROMPT "<prompt_name>" ANSWER "<prompt_answer>" [, PROMPT "<prompt_name>" ANSWER "<prompt_answer>" ...]];

		1st Example of MicroStrategy Command Manager "CREATE MOBILESUBSCRIPTION" syntax implemented
		CREATE MOBILESUBSCRIPTION "My first mobile subscription" SCHEDULE "First of Month" USER "architect" CONTENT "Customer Income Analysis" IN FOLDER  "\Public Objects\REPORTS\SUBJECT Areas\Customer Analysis" IN PROJECT "MicroStrategy Tutorial" RUNFRESH TRUE MODIFICATIONBYRECIPIENTS FALSE;
		
		
#>
       
[CmdletBinding()]
param(
[Parameter(Mandatory=$True,ParameterSetName='HistoryList')]
[Switch] $HistoryListSubscription, 
[Parameter(Mandatory=$True,ParameterSetName='Email')]
[Switch] $EmailSubscription, 
[Parameter(Mandatory=$True,ParameterSetName='File')]
[Switch] $FileSubscription, 
[Parameter(Mandatory=$True,ParameterSetName='Print')]
[Switch] $PrintSubscription, 
[Parameter(Mandatory=$True,ParameterSetName='Cache')]
[Switch] $CacheUpdateSubscription, 
[Parameter(Mandatory=$True,ParameterSetName='Mobile')]
[Switch] $MobileSubscription, 
[string]$SubscriptionName,
[string]$OwnerLogin,
[Parameter(Mandatory=$true)]
[string]$ScheduleName,
[string]$UserLogin,
[string]$UserGroup,
[string]$AddressName,
[string]$ContactName,
[string]$ContactAddressName,
[string]$ContactGroupName,
[string]$ContentName,
[string]$ContentPath,
[string]$ContentGUID,
[Parameter(Mandatory=$true)]
[ValidateSet("REPORT","DOCUMENT")]
[string]$ContentType,
[Parameter(Mandatory=$true)]
[string]$ProjectName,
[Parameter(ParameterSetName='Print')]
[Switch]$PrintAll,
[Parameter(ParameterSetName='Print')]
[Switch]$PrintCollate,
[Parameter(ParameterSetName='Print')]
[int[]]$PrintRange,
[Parameter(ParameterSetName='Print')]
[int]$PrintCopies, 
[string]$Subject,
[Parameter(ParameterSetName='Cache')]
[ValidateSet("HTML", "EXCEL", "PDF", "FLASH")] 
[string]$CacheUpdateDeliveryFormat,
[Parameter(ParameterSetName='Email')]
[Parameter(ParameterSetName='File')]
[ValidateSet("HTML", "CSV", "PLAINTEXT", "EXCEL", "PDF", "FLASH")] 
[string]$DeliveryFormat,
[string]$FileName,
[Parameter(ParameterSetName='Email')]
[Parameter(ParameterSetName='File')]
[ValidateSet("COMMA", "SPACE", "TAB", "SEMICOLON", "COLON", "OTHER")]
[string]$TextDelimiter,
[Parameter(ParameterSetName='Email')]
[Parameter(ParameterSetName='File')]
[string]$OtherTextDelimiter,
[Parameter(ParameterSetName='Email')]
[Parameter(ParameterSetName='File')]
[string]$ZipFileName,
[string]$ZipFilePassword,
[string]$Message,
[DateTime]$ExpirationDate,
[Switch]$SendToHistoryList,
[Switch]$SendNotification,
[string]$SendNotificationAddress,
[string]$PhysicalAddress,
[string]$DeviceName,
[Switch]$RunFresh,
[Switch]$CreateUpdateCache,
[Switch]$IncludeLink,
[Switch]$IncludeData,
[Switch]$OverWriteOldVersion,
[string]$HistoryLinkUrl,
[Parameter(ParameterSetName='HistoryList')]
[Parameter(ParameterSetName='Email')]
[string]$HistoryLinkWebServer,
[Switch]$SendPreviewNow,
[Switch]$ModificationByRecipients,
[string[]]$Prompt,
[string[]]$PromptAnswer
)

$CmdLine = $null
$Empty = [String]::Empty
switch ($true) { 
  $HistoryListSubscription          		{ $CmdLine += "CREATE HISTORYLISTSUBSCRIPTION "}                                       
  $EmailSubscription                		{ $CmdLine += "CREATE EMAILSUBSCRIPTION "}                                            
  $FileSubscription                 		{ $CmdLine += "CREATE FILESUBSCRIPTION "}                                               
  $PrintSubscription                		{ $CmdLine += "CREATE PRINTSUBSCRIPTION "}                                              
  $CacheUpdateSubscription          		{ $CmdLine += "CREATE CACHEUPDATESUBSCRIPTION "}                                     
  $MobileSubscription	            		{ $CmdLine += "CREATE MOBILESUBSCRIPTION "}
  ($SubscriptionName -ne $Empty)     		{ $CmdLine += "`"$SubscriptionName`" " }
  ($OwnerLogin -ne $Empty)					{ $CmdLine += "FOR OWNER `"$OwnerLogin`" " }
  ($ScheduleName -ne $Empty)		   		{ $CmdLine += "SCHEDULE `"$ScheduleName`" " }
  ($AddressName -ne $Empty)					{ $CmdLine += "ADDRESS `"$AddressName`" " }
  ($UserLogin -ne $Empty) 					{ $CmdLine += "USER `"$UserLogin`" " }
  ($UserGroup  -ne $Empty)					{ $CmdLine += "GROUP `"$UserGroup`" " }
  ($ContactName -ne $Empty)					{ $CmdLine += "CONTACT `"$ContactName`" " }
  ($ContactAddressName -ne $Empty)			{ $CmdLine += "ADDRESS `"$ContactAddressName`" " }
  ($ContactGroupName -ne $Empty)			{ $CmdLine += "CONTACTGROUP `"$ContactGroupName`" " }
  ($ContentName -ne $Empty)					{ $CmdLine += "CONTENT `"$ContentName`" IN FOLDER `"$ContentPath`" "} 
  ($ContentGUID -ne $Empty)					{ $CmdLine += "GUID $ContentGUID "}
  ($ContentType  -ne $Empty) 				{ $CmdLine += "CONTENTTYPE $ContentType " }
  ($ProjectName  -ne $Empty)				{ $CmdLine += "IN PROJECT `"$ProjectName`" "}
  ($Subject  -ne $Empty) 					{ $CmdLine += "SUBJECT `"$Subject`" " }
  ($CacheUpdateDeliveryFormat -ne $Empty)  	{ $CmdLine += "DELIVERYFORMAT `"$CacheUpdateDeliveryFormat`" " }
  ($DeliveryFormat -ne $Empty) 				{ $CmdLine += "DELIVERYFORMAT $DeliveryFormat " }
  ($FileName -ne $Empty)         			{ $CmdLine += "FILENAME `"$FileName`" "}
  ($TextDelimiter -ne $Empty)               { $CmdLine += "TEXTDELIMITER `"$TextDelimiter`" "}
  ($OtherTextDelimiter -ne $Empty)          { $CmdLine += " `"$OtherTextDelimiter`" " }
  ($ZipFileName -ne $Empty)                 { $CmdLine += "COMPRESSCONTENTS ZIPFILENAME `"$ZipFileName`" "}
  ($ZipFilePassword -ne $Empty)             { $CmdLine += "ZIPFILEPASSWORD `"$ZipFilePassword`" "}
  ($Message -ne $Empty)                     { $CmdLine += "MESSAGE `"$Message`" "}
  $OverWriteOldVersion              		{ $CmdLine += "OVERWRITEOLDVERSION TRUE " }
  ($HistoryLinkUrl -ne $Empty)         		{ $CmdLine += "HISTORYLINKURL `"$HistoryLinkUrl`" " }
  ($HistoryLinkWebServer -ne $Empty)        { $CmdLine += "HISTORYLINKWEBSERVER `"$HistoryLinkWebServer`" " }
  ($ExpirationDate -match "\d+/\d+/\d+")              { $CmdLine += "EXPIRATIONDATE $ExpirationDate " }
  $SendToHistoryList           		   { $CmdLine += "SENDTOHISTORYLIST TRUE " }
  $IncludeLink                              { $CmdLine += "INCLUDELINK TRUE " }
  $SendPreviewNow                           { $CmdLine += "SENDPREVIEWNOW TRUE " }
  $IncludeData                      		{ $CmdLine += "INCLUDEDATA TRUE " }
  $SendNotification                 		{ $CmdLine += "SENDNOTIFICATION TRUE " }
  ($SendNotificationAddress -ne $Empty)	    { $CmdLine += "SENDNOTIFICATION_ADDRESS `"$SendNotificationAddress`" " }
  ($PhysicalAddress -ne $Empty)	            { $CmdLine += "PHYSICALADDRESS `"$PhysicalAddress`" " }
  ($DeviceName -ne $Empty)                  { $CmdLine += "DEVICE `"$DeviceName`" " }
  $RunFresh      			                { $CmdLine += "RUNFRESH TRUE " }
  $CreateUpdateCache        		        { $CmdLine += "CREATEUPDATECACHE TRUE " }
  $ModificationByRecipients         		{ $CmdLine += "MODIFICATIONBYRECIPIENTS TRUE " }
 }
 
if($Prompt) { 
   $Loops = $Prompt.Length - 1
   for ($c = 0; $c -le $Loops; $c++) {
		$p = $Prompt[$c];
		$a = $PromptAnswer[$c];
		$CmdLine += "PROMPT `"$p`" ANSWER `"$a`" "
		if($Loops -gt 0) { $CmdLine += ", " } 
		}
	$CmdLine.TrimEnd(", ")  
}
$CmdLine += ";"
$CmdLine
}