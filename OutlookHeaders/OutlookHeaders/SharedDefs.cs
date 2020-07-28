//  
//    OutlookHeaders Add-in
//    "Outlook add-in for modifying mail headers & outbound emails."
//    Copyright (c) 2020 www.dennisbabkin.com
//    
//        https://dennisbabkin.com/olh
//    
//    Licensed under the Apache License, Version 2.0 (the "License");
//    you may not use this file except in compliance with the License.
//    You may obtain a copy of the License at
//    
//        https://www.apache.org/licenses/LICENSE-2.0
//    
//    Unless required by applicable law or agreed to in writing, software
//    distributed under the License is distributed on an "AS IS" BASIS,
//    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//    See the License for the specific language governing permissions and
//    limitations under the License.
//  
//



//These are definitions that are share among multiple projects (including installers)

namespace GlobalDefs
{
	class GlobDefs
	{
		public const string gkstrAppName = "OutlookHeaders";
		public const string gkstrAppNameFull = gkstrAppName + " Add-in";
		public const string gkstrEventSrcName = gkstrAppNameFull;

		public const string gkstrAppRegCompany = "www.dennisbabkin.com";									//Name of the company in system registry that made this add-in
		public const string gkstrAppRegistryCompany = @"Software\" + gkstrAppRegCompany;					//Registry key for this company
		public const string gkstrAppRegistryCompanyWow64 = @"Software\Wow6432Node\" + gkstrAppRegCompany;	//Same as 'gkstrAppRegistryCompany' but when running on a 64-bit system as 32-bit module
		public const string gkstrAppRegistryApp = gkstrAppRegistryCompany + @"\" + gkstrAppName;			//Registry key for this add-in
		public const string gkstrAppRegistryKey = gkstrAppRegistryApp + @"\Settings";						//Registry key for this add-in private settings
		public const string gkstrAppRegVal_InstallFldr = "path";											//Folder path to where add-in is installed
		public const string gkstrAppRegVal_Version = "ver";													//Version of this add-in that is installed
		public const string gkstrAppRegVal_ConfigFilePaths = "cfgs";										//All file paths to all user.config files used by all versions

		public const string gkstrAppRegVal_1stRun_ImportDataFile = "{A23C1358-9F5F-4011-BA7F-A31D6982DFE5}";	//[Always saved in HKLM key] Path to the import-file to import during the first run
		public const string gkstrAppRegVal_1stRun_StartupFlags = "{D884C555-864A-4c73-961E-9ED626A0954C}";		//[Always saved in HKLM key] Value of flags to use during the first run. Bitwise:
																												//	1 = to start this add-in disabled
																												//	2 = To log debugging info into diagnostic event log when emails are sent out

		public const string gkstrExportFileExt = "ohxml";													//Do not use a dot! File extension for import/export configuration files

	}
}

