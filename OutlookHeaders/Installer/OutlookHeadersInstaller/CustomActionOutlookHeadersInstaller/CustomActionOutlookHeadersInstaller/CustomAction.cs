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



using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Deployment.WindowsInstaller;

using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Security.Principal;
using System.IO;
using Microsoft.Win32;
using System.Reflection;



namespace CustomActionOutlookHeadersInstaller
{
	public class CustomActions
	{
		public static CurrentVersion gAppVersion = new CurrentVersion();			//Will be added later from the assembly file

		public class CurrentVersion
		{
			public string getVersion()
			{
				//RETURN: = This app version as string
				if (_strAppVer == null)
				{
					Assembly asm = Assembly.GetExecutingAssembly();
					_strAppVer = asm.GetName().Version.ToString();
				}

				return _strAppVer;
			}

			private string _strAppVer = null;
		}




		//[DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
		//public static extern int MessageBox(int hWnd, String text, String caption, uint type);


		[CustomAction]
		public static ActionResult caFirstStage(Session session)
		{
			//Called during the first stage of installation, uninstallation, change, repair or upgrade (or BEFORE it)

			//What stage are we on
			MSI_INFO msiInf = determineStage(session, false);
			if (msiInf.stage == MSI_STAGE.MS_Unknown)
			{
				//Error
				LogDiagnosticSpecError(1058, session);
				return ActionResult.Failure;
			}


			//Process specific stages
			if (msiInf.stage == MSI_STAGE.MS_INSTALL ||
				msiInf.stage == MSI_STAGE.MS_REPAIR)
			{
				//Register event log source
				registerEventSource();
			}


			if (msiInf.stage == MSI_STAGE.MS_INSTALL)
			{
				//Log event that we're installing for the first time
				LogDiagnosticEvent(1056, "First installation:\n" +
					"Path: " + msiInf.strInstallFolder + "\n" +
					"MSI: " + msiInf.strMSIFolder + "\n" +
					"RegistryKey: " + msiInf.strRegistryKey + "\n" +
					"User: " + msiInf.strUserName + "\n" +
					"Bundled: " + msiInf.strBundledInstall + "\n" +
					"ImportData: " + msiInf.strImportDataPath + "\n" +
					"Flags: " + msiInf.strStartFlags + "\n"
				, EventLogEntryType.Information);
			}
			else
			{
				//Log other event
				LogDiagnosticEvent(1057, "Before: " + msiInf.stage.ToString() + "\n" +
					"MSI: " + msiInf.strMSIFolder + "\n" +
					"User: " + msiInf.strUserName + "\n" +
					"Bundled: " + msiInf.strBundledInstall + "\n"
					,
					EventLogEntryType.Information);
			}



			return ActionResult.Success;
		}




		[CustomAction]
		public static ActionResult caLastStage(Session session)
		{
			//Called during the last stage of installation, uninstallation, change, repair or upgrade

			//What stage are we on
			MSI_INFO msiInf = determineStage(session, true);
			if (msiInf.stage == MSI_STAGE.MS_Unknown)
			{
				//Error
				LogDiagnosticSpecError(1059, session);
				return ActionResult.Failure;
			}

			if (msiInf.stage == MSI_STAGE.MS_INSTALL)
			{
				//Finishing first installation
				processInstallationParams(msiInf);

				//Set message
				LogDiagnosticEvent(1061, "After: " + msiInf.stage.ToString() + "\n" +
					"MSI: " + msiInf.strMSIFolder + "\n" +
					"User: " + msiInf.strUserName + "\n" +
					"Bundled: " + msiInf.strBundledInstall + "\n"
					,
					EventLogEntryType.Information);
			}
			else if (msiInf.stage == MSI_STAGE.MS_UNINSTALL)
			{
				//Uninstallation
				LogDiagnosticEvent(1060, "Uninstallation:\n" +
					"Path: " + msiInf.strInstallFolder + "\n" +
					"MSI: " + msiInf.strMSIFolder + "\n" +
					"RegistryKey: " + msiInf.strRegistryKey + "\n" +
					"User: " + msiInf.strUserName + "\n" +
					"Bundled: " + msiInf.strBundledInstall + "\n"
					, EventLogEntryType.Information);

				//Remove user settings (depending on installation type)
				removeAllUserSettings(msiInf);


				//Unregister event source
				unregisterEventSource();
			}
			else
			{
				//Log other event
				LogDiagnosticEvent(1061, "After: " + msiInf.stage.ToString() + "\n" +
					"MSI: " + msiInf.strMSIFolder + "\n" +
					"User: " + msiInf.strUserName + "\n" +
					"Bundled: " + msiInf.strBundledInstall + "\n"
					,
					EventLogEntryType.Information);
			}


			return ActionResult.Success;
		}


		private static void processInstallationParams(MSI_INFO msiInf)
		{
			//Process installation parameters during first installation
			string strImportDataFilePath = null;
			string strStartupFlags = null;


			//Get folder path to the same location where this MSI was running from
			string strMsiFldr = "";
			if (!string.IsNullOrEmpty(msiInf.strMSIFolder))
			{
				strMsiFldr = makeFolderEndWithSlash(msiInf.strMSIFolder, true);
			}
			else
				LogDiagnosticSpecWarning(1089);		//We should've gotten the folder here!


			//See if we have an import data file?
			if (!string.IsNullOrEmpty(msiInf.strImportDataPath))
			{
				//Is it a file?
				if (File.Exists(msiInf.strImportDataPath))
				{
					//Yes! Then use it
					strImportDataFilePath = msiInf.strImportDataPath;
				}
			}

			if (string.IsNullOrEmpty(strImportDataFilePath))
			{
				//Additionally, see if there's a special file placed in the same folder as the MSI
				if (!string.IsNullOrEmpty(strMsiFldr))
				{
					string strPath = strMsiFldr + "ImportData." + GlobalDefs.GlobDefs.gkstrExportFileExt;

					//See if this file exists
					if (File.Exists(strPath))
					{
						//Yes, then use it!
						strImportDataFilePath = strPath;
					}
				}
			}


			//Do we have startup flags?
			if (!string.IsNullOrEmpty(msiInf.strStartFlags))
			{
				//Use it
				strStartupFlags = msiInf.strStartFlags;
			}

			if (string.IsNullOrEmpty(strStartupFlags))
			{
				//Additionally, see if there's a special file placed in the same folder as MSI
				if (!string.IsNullOrEmpty(strMsiFldr))
				{
					string strPath = strMsiFldr + "StartupFlags.txt";

					//See if this file exists
					if (!File.Exists(strPath))
					{
						//Then try without an extension
						strPath = strMsiFldr + "StartupFlags";
						if (!File.Exists(strPath))
						{
							strPath = null;
						}
					}

					if (!string.IsNullOrEmpty(strPath))
					{
						try
						{
							//Read this file
							string strTxt = System.IO.File.ReadAllText(strPath).Trim();

							//This is our value - convert it to integer
							if (!string.IsNullOrEmpty(strTxt))
							{
								//Got it! Then use it
								strStartupFlags = strTxt;
							}
						}
						catch (Exception ex)
						{
							//Failed
							LogDiagnosticSpecError(1090, "path: " + strPath + " > " + ex.ToString());
						}
					}
				}
			}


			//Convert to flags
			UInt32 uiFlags = 0;
			bool bHaveFlags = UInt32.TryParse(strStartupFlags, out uiFlags);

			if(bHaveFlags || !string.IsNullOrEmpty(strImportDataFilePath))
			{
				//Now save these values in the registry where they will be read by the add-in during its first run.
				//INFO: These always go into the HKLM hive.
				try
				{
					using (RegistryKey rK = Registry.LocalMachine.CreateSubKey(GlobalDefs.GlobDefs.gkstrAppRegistryKey))
					{
						string strLogMsg = "";

						if(!string.IsNullOrEmpty(strImportDataFilePath))
						{
							//Save in registry
							rK.SetValue(GlobalDefs.GlobDefs.gkstrAppRegVal_1stRun_ImportDataFile, strImportDataFilePath, RegistryValueKind.ExpandString);

							if (!string.IsNullOrEmpty(strLogMsg))
								strLogMsg += "; ";

							strLogMsg += "ImportDataFrom: " + strImportDataFilePath;
						}

						if (bHaveFlags)
						{
							//Save in registry
							rK.SetValue(GlobalDefs.GlobDefs.gkstrAppRegVal_1stRun_StartupFlags, uiFlags, RegistryValueKind.DWord);

							if (!string.IsNullOrEmpty(strLogMsg))
								strLogMsg += "; ";

							strLogMsg += "StartupFlags: " + uiFlags;
						}

						if (!string.IsNullOrEmpty(strLogMsg))
						{
							//Log it
							LogDiagnosticEvent(1127, "Found initial install params: " + strLogMsg, EventLogEntryType.Information);
						}
					}
				}
				catch (Exception ex)
				{
					//Failed
					LogDiagnosticSpecError(1091, ex);
				}
			}
		}


		private static void removeAllUserSettings(MSI_INFO msiInf)
		{
			//Remove all user settings during final uninstallation
			//INFO: This includes all user.config files and system registry settings


			//See if we're doing ALL-USERs install
			if (!string.IsNullOrEmpty(msiInf.strUserName))
			{
				//Specific user
				LogDiagnosticEvent(1117, "Removing user settings only for user: " + msiInf.strUserName, EventLogEntryType.Information);

				if (!removeSpecUserSettings(msiInf.strUserName))
				{
					//Failed to remove everything for this user
					LogDiagnosticSpecError(1109, "Failed to remove some (or all) user settings: " + msiInf.strUserName);
				}

			}
			else
			{
				//All users
				LogDiagnosticEvent(1118, "Removing user settings for all users on this computer.", EventLogEntryType.Information);

				//Enumerate all available user accounts on this system
				//INFO: This is not the best way to do it since if the user isn't logged in, their registry profile will not be loaded!

				try
				{
					using (RegistryKey regUsrsK = Registry.Users.OpenSubKey("", false))
					{
						string[] strUsrSids = regUsrsK.GetSubKeyNames();

						List<string> arrUsrSids2Del = new List<string>();

						foreach (string strUsrSID in strUsrSids)
						{
							//See if this user has our software installed
							using (RegistryKey regK = regUsrsK.OpenSubKey(strUsrSID + @"\" + GlobalDefs.GlobDefs.gkstrAppRegistryApp, false))
							{
								if (regK != null)
								{
									//Delete it for this user (later)
									arrUsrSids2Del.Add(strUsrSID);
								}
							}
						}


						//And remove each user data
						foreach (string strSID in arrUsrSids2Del)
						{
							//Add log entry
							LogDiagnosticEvent(1119, "Removing user settings for user SID: " + strSID, EventLogEntryType.Information);

							if (!removeSpecUserSettingsBySID(strSID))
							{
								LogDiagnosticSpecError(1116, "Failed to remove some (or all) user settings for SID: " + strSID);
							}
						}
					}
				}
				catch (Exception ex)
				{
					//Failed
					LogDiagnosticSpecError(1113, ex);
				}

			}


			//Then finally remove main HLKM key for the app (that may be present)
			try
			{
				string strAppKey = GlobalDefs.GlobDefs.gkstrAppRegistryApp;
				bool bOkDel = false;
				using (RegistryKey regK = Registry.LocalMachine.OpenSubKey(strAppKey, false))
				{
					if (regK != null)
						bOkDel = true;
				}

				if (bOkDel)
				{
					//Delete with subkeys
					Registry.LocalMachine.DeleteSubKeyTree(strAppKey);

					//Log
					LogDiagnosticEvent(1125, @"Removed reg-key: HTML\" + strAppKey, EventLogEntryType.Information);
				}
			}
			catch (Exception ex)
			{
				//Error
				LogDiagnosticSpecError(1111, ex);
			}


			//And delete company key (but only if it was not empty)
			//Ex: HKLM\Software\www.dennisbabkin.com
			string strErrDesc;
			string strCompKey = GlobalDefs.GlobDefs.gkstrAppRegistryCompany;
			if (deleteRegKeyIfEmpty(Registry.LocalMachine, strCompKey, out strErrDesc))
			{
				//Log
				LogDiagnosticEvent(1126, @"Removed reg-key: HTML\" + strCompKey, EventLogEntryType.Information);
			}
			else
			{
				//Error
				LogDiagnosticSpecError(1112, strErrDesc);
			}

		}



		private static bool removeSpecUserSettings(string strUserName)
		{
			//Remove settings for user with name 'strUserName'
			//RETURN:
			//		= true if no errors

			//Convert name to SID
			string strUserSID = getSIDfromUserName(strUserName);
			if (string.IsNullOrEmpty(strUserSID))
			{
				LogDiagnosticSpecError(1101, "usr=" + strUserName);
				return false;
			}

			return removeSpecUserSettingsBySID(strUserSID);
		}


		private static bool removeSpecUserSettingsBySID(string strUserSID)
		{
			//Remove settings for user with 'strUserSID', ex: "S-1-5-21-664853314-724024733-1308763421-1001"
			//RETURN:
			//		= true if no errors
			bool bRes = true;

			//Get user name for the SID (need it for debugging output only)
			string strUserName = getUserNameFromSID(strUserSID);


			string strKey = strUserSID + @"\" + GlobalDefs.GlobDefs.gkstrAppRegistryKey;

			try
			{
				//Open HKEY_USERS\<SID>\Software\www.dennisbabkin.com\OutlookHeaders\Settings
				using (RegistryKey regK = Registry.Users.OpenSubKey(strKey, true))
				{
					if (regK != null)
					{
						int nOSError;

						//Get paths of the config files
						string[] arrCFPs = (string[])regK.GetValue(GlobalDefs.GlobDefs.gkstrAppRegVal_ConfigFilePaths, new string[] { }, RegistryValueOptions.None);

						//Go through all and delete them
						foreach (string strCfgPath in arrCFPs)
						{
							try
							{
								//Ex: C:\Users\Admin\AppData\Local\Microsoft_Corporation\OutlookHeaders.vsto_vstol_Path_i5vznyicnelzuytw513coiw3qpofvydj\12.0.4518.1014\user.config
								File.Delete(strCfgPath);

								LogDiagnosticEvent(1120, "Removed cfg: " + strCfgPath, EventLogEntryType.Information);
							}
							catch (Exception ex)
							{
								//Failed to delete
								LogDiagnosticSpecError(1104, "Usr=" + strUserName + ", SID=" + strUserSID + " File=\"" + strCfgPath + "\" > " + ex.ToString());
								bRes = false;
							}

							//Go one level up
							string strParDir = Path.GetDirectoryName(strCfgPath);
							if (!string.IsNullOrEmpty(strParDir))
							{
								//Delete folder if not empty
								//Ex: C:\Users\Admin\AppData\Local\Microsoft_Corporation\OutlookHeaders.vsto_vstol_Path_i5vznyicnelzuytw513coiw3qpofvydj\12.0.4518.1014
								if (DeleteDirectoryIfNotEmpty(strParDir, out nOSError))
								{
									//Log it
									LogDiagnosticEvent(1121, "Removed fldr: " + strParDir, EventLogEntryType.Information);
								}
								else
								{
									if (nOSError != 145)		//ERROR_DIR_NOT_EMPTY
									{
										//Failed to delete because of some other reason than directory-not-empty
										LogDiagnosticSpecError(1105, "Usr=" + strUserName + ", SID=" + strUserSID + " Dir=\"" + strParDir + "\" OSError=" + nOSError);
										bRes = false;
									}
								}
							}

							//Go another level up
							strParDir = Path.GetDirectoryName(strParDir);
							if (!string.IsNullOrEmpty(strParDir))
							{
								//Delete it if not empty
								//Ex: C:\Users\Admin\AppData\Local\Microsoft_Corporation\OutlookHeaders.vsto_vstol_Path_i5vznyicnelzuytw513coiw3qpofvydj
								if (DeleteDirectoryIfNotEmpty(strParDir, out nOSError))
								{
									//Log it
									LogDiagnosticEvent(1122, "Removed fldr: " + strParDir, EventLogEntryType.Information);
								}
								else
								{
									if (nOSError != 145)		//ERROR_DIR_NOT_EMPTY
									{
										//Failed to delete because of some other reason than directory-not-empty
										LogDiagnosticSpecError(1106, "Usr=" + strUserName + ", SID=" + strUserSID + " Dir=\"" + strParDir + "\" OSError=" + nOSError);
										bRes = false;
									}
								}
							}


						}

					}
					else
					{
						//No key?
						//LogDiagnosticSpecError(1103, "Usr=" + strUserName + ", SID=" + strUserSID + " Key=" + strKey);
						//bRes = false;
					}
				}
			}
			catch (Exception ex)
			{
				//Failed
				LogDiagnosticSpecError(1102, "Usr=" + strUserName + ", SID=" + strUserSID + " > " + ex.ToString());
				bRes = false;
			}



			//Then delete the registry key for the app
			//Ex: HKEY_USERS\<SID>\Software\www.dennisbabkin.com\OutlookHeaders
			string strAppRegKey = strUserSID + @"\" + GlobalDefs.GlobDefs.gkstrAppRegistryApp;

			try
			{
				//Delete key of it exists
				bool bDelIt = false;
				using (RegistryKey regK = Registry.Users.OpenSubKey(strAppRegKey))
				{
					if (regK != null)
						bDelIt = true;
				}

				if (bDelIt)
				{
					//Remove it with all subkeys!
					Registry.Users.DeleteSubKeyTree(strAppRegKey);

					//Log it
					LogDiagnosticEvent(1123, "Removed reg-key: " + strAppRegKey, EventLogEntryType.Information);
				}
			}
			catch (Exception ex)
			{
				//Failed
				LogDiagnosticSpecError(1107, "Usr=" + strUserName + ", SID=" + strUserSID + ", DelKey=" + strAppRegKey + " > " + ex.ToString());
				bRes = false;
			}


			//And delete company key (but only if it was not empty)
			//Ex: HKEY_USERS\<SID>\Software\www.dennisbabkin.com
			string strErrDesc;
			string strCompKey = strUserSID + @"\" + GlobalDefs.GlobDefs.gkstrAppRegistryCompany;
			if (deleteRegKeyIfEmpty(Registry.Users, strCompKey, out strErrDesc))
			{
				//Log it
				LogDiagnosticEvent(1124, "Removed reg-key: " + strCompKey, EventLogEntryType.Information);
			}
			else
			{
				LogDiagnosticSpecError(1108, "Usr=" + strUserName + ", SID=" + strUserSID + " > " + strErrDesc);
				bRes = false;
			}


			return bRes;
		}


		private static bool deleteRegKeyIfEmpty(RegistryKey regHive, string strKey, out string strErrDesc)
		{
			//Delete registry key from 'strKey' only if it has no keys or values
			//WARNING: This function is not thread-safe!
			//'regHive' = registry hive for 'strKey'
			//'strErrDesc' = receives error description
			//RETURN:
			//		= true if success
			bool bRes = false;
			strErrDesc = "";

			try
			{
				bool bOK2Del = false;
				using (RegistryKey regK = regHive.OpenSubKey(strKey, true))
				{
					//Only if key existed
					if (regK != null)
					{
						if (regK.SubKeyCount == 0 && regK.ValueCount == 0)
						{
							//Can delete it
							bOK2Del = true;
						}
					}
				}

				if (bOK2Del)
				{
					//Delete this key
					regHive.DeleteSubKey(strKey, false);
				}

				//Done
				bRes = true;
			}
			catch (Exception ex)
			{
				//Failed
				strErrDesc = "[1110] Hive=" + regHive.ToString() + " Key=" + strKey + " > " + ex.ToString();
				bRes = false;
			}

			return bRes;
		}

		private static string getSIDfromUserName(string strUserName)
		{
			//RETURN:
			//		= SID for a user name, ex: "S-1-5-21-664853314-724024733-1308763421-1001", or
			//		= null if error
			try
			{
				if (string.IsNullOrEmpty(strUserName))
				{
					//User name must be provided
					throw new Exception("[1100]");
				}

				NTAccount ntAct = new NTAccount(strUserName);
				SecurityIdentifier sid = (SecurityIdentifier)ntAct.Translate(typeof(SecurityIdentifier));
				return sid.ToString();
			}
			catch(Exception ex)
			{
				//Exception
				LogDiagnosticSpecError(1099, "Usr: \"" + strUserName + "\" > " + ex.ToString());
				return null;
			}
		}

		private static string getUserNameFromSID(string strSID)
		{
			//RETURN:
			//		= User name for 'strSID', ex: "COMPUTER\User"
			//		= null if error
			try
			{
				if (string.IsNullOrEmpty(strSID))
				{
					//User name must be provided
					throw new Exception("[1115]");
				}

				SecurityIdentifier secId = new SecurityIdentifier(strSID);
				NTAccount ntAct = (NTAccount)secId.Translate(typeof(NTAccount));
				return ntAct.ToString();
			}
			catch (Exception ex)
			{
				//Exception
				LogDiagnosticSpecError(1114, "SID: \"" + strSID + "\" > " + ex.ToString());
				return null;
			}

		}


		[DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
		static extern bool RemoveDirectory(string lpPathName);

		public static bool DeleteDirectoryIfNotEmpty(string strPath, out int nOutWin32Error)
		{
			bool bRes = RemoveDirectory(strPath);
			int nErr = Marshal.GetLastWin32Error();
			nOutWin32Error = !bRes ? nErr : 0;
			return bRes;
		}


		private enum MSI_STAGE
		{
			MS_Unknown,
			MS_INSTALL,				//Called when installing new app (when it was never installed before)
									//IMPORTANT: If 'bAfter' is true, registry keys/values that were marked as HKCU will not be set here, if installer is running as system!
			MS_UNINSTALL,			//Called to completely remove the app (when user chose to uninstall it)
			MS_CHANGE,				//Called to change installation features
			MS_REPAIR,				//Called to repair existing installation (files, registry settings, etc.)
			MS_UPGRADE_PREV,		//Called by the previous version of MSI when upgrading from it to the new version of MSI (it is called first)
			MS_UPGRADE_NEW,			//Called by the new version of MSI when upgrading to it from an old version of MSO (it is called second)
									//IMPORTANT: The sequence is executes as follows:
									//				MS_UPGRADE_PREV,	'bAfter' = false	=> installation files and register are still from previous version
									//				MS_UPGRADE_PREV,	'bAfter' = true		=> installation files may've been deleted (if there were changes between versions), and registry keys may've been deleted (if there were changes between versions). Note that if there's no changes - they will not be deleted here!
									//				MS_UPGRADE_NEW,		'bAfter' = false	=> no change since last stage
									//				MS_UPGRADE_NEW,		'bAfter' = true		=> installed files and registry keys have been updated
		}

		private class MSI_INFO
		{
			public MSI_STAGE stage = MSI_STAGE.MS_Unknown;		//MSI installer stage

			public WindowsIdentity caller = null;				//What security identify we're running under, or null if known
			public string strBundledInstall = "";				//"1" for Bundled installation
			public string strUserName = "";						//User name that we're installing for, or "" if doing it for all-users
			public string strInstallFolder = "";				//Folder where the app is being installed to (always has no terminating slash)
			public string strMSIFolder = "";					//Folder where MSI package is running from (always has no terminating slash)
			public string strRegistryKey = "";					//Registry key for this app (always has no terminating slash)
			public string strImportDataPath = "";				//Provided by end-user from a command line call to MSIEXEC - file with XML data to import during installation
			public string strStartFlags = "";					//String with bitwise flag values, provided by end-user from a command line call to MSIEXEC:
																//	INFO: see 'gkstrAppRegVal_1stRun_StartupFlags' for info
		}

		private enum IDX_PART
		{
			IDX_P_INSTALLED,
			IDX_P_REINSTALL,
			IDX_P_UPGRADEPRODUCTCODE,
			IDX_P_REMOVE,

			IDX_P_INSTALLFOLDER,
			IDX_P_REG_KEY,
			IDX_P_USER_NAME,
			IDX_P_BUNDLED_INSTALLER,
			IDX_P_SOURCE_DIR,

			IDX_P_IMPORT_DATA,
			IDX_P_STARTUP_FLAGS,

			IDX_P_Count						//Must be last
		}


		private static MSI_INFO determineStage(Session session, bool bAfter)
		{
			//'bAfter' = true if called after the stage, false - if before
			//RETURN:
			//		= MSI stage info
			MSI_INFO res = new MSI_INFO();

			int nCnt = session.CustomActionData.Count;
			if (nCnt > 0)
			{
				string strCustData = session.CustomActionData.ToString();
				string[] arrParts = strCustData.Split(new string[] { "|||" }, StringSplitOptions.None);

				int nCntPrts = arrParts.Length;
				if (nCntPrts == (int)IDX_PART.IDX_P_Count)
				{
					//Get passed paths
					res.strInstallFolder = makeFolderEndWithSlash(arrParts[(int)IDX_PART.IDX_P_INSTALLFOLDER].Trim(), false);
					res.strMSIFolder = makeFolderEndWithSlash(arrParts[(int)IDX_PART.IDX_P_SOURCE_DIR].Trim(), false);
					res.strRegistryKey = makeFolderEndWithSlash(arrParts[(int)IDX_PART.IDX_P_REG_KEY].Trim(), false);
					res.strUserName = arrParts[(int)IDX_PART.IDX_P_USER_NAME].Trim();
					res.strBundledInstall = arrParts[(int)IDX_PART.IDX_P_BUNDLED_INSTALLER].Trim();
					res.strImportDataPath = makeFolderEndWithSlash(arrParts[(int)IDX_PART.IDX_P_IMPORT_DATA].Trim(), false);

					//And get flags
					res.strStartFlags = arrParts[(int)IDX_PART.IDX_P_STARTUP_FLAGS].Trim();

					//Check what account we're running under
					res.caller = WindowsIdentity.GetCurrent();


					const string kStrRegVal_InstallFolderExisted = "{0AEA8881-6418-4584-8C20-D4BDED8BAE04}";


					//SOURCE:
					//		https://stackoverflow.com/a/17608049/843732

					//Tested the following experimentally with this particular installer:
					//	[0] = Installed
					//	[1] = REINSTALL
					//	[2] = UPGRADINGPRODUCTCODE
					//	[3] = REMOVE
					//	[4] = INSTALLFOLDER
					//	[5] = Registry key
					//	[6] = User Name
					//	[7] = Bundled installer
					//	[8] = Source directory
					//	[9] = IMPORTDATA (provided from command line)
					//	[10] = STARTUPFLAGS (provided from command line)
					//
					//									[0]			[1]			[2]			[3]			[4]						[9][10]					Runs-as:
					//FRESH INSTALL:					-			-			-			-			InstallFldrPath			Yes						S-1-5-18 = Local System = (Shows UAC prompt)
					//
					//UNINSTALL:						Str			-			-			Str			InstallFldrPath			-						S-1-5-18 = Local System = (Shows UAC prompt)
					//(From Control Panel or via
					// context-menu "Uninstall" command
					// in Windows Explorer)
					//
					//Repeat Install:					Str			-			-			-			InstallFldrPath			- , or					S-1-5-18 = Local System = (DOES NOT SHOW UAC PROMPT!)
					//(When MSI file is double-																					Yes (if started with
					// clicked when that same																					cmd line params from
					// version is already																						msiexec)
					// installed)
					//
					//CHANGE:							Str			-			-			-			InstallFldrPath			-						S-1-5-18 = Local System = (DOES NOT SHOW UAC PROMPT!)
					//(From Control Panel)
					//
					//REPAIR:							Str			Str			-			-			InstallFldrPath			-						S-1-5-18 = Local System = (Shows UAC prompt)
					//(From Control Panel)
					//
					//REPAIR:							Str			Str			-			-			InstallFldrPath			-						S-1-5-18 = Local System = (DOES NOT SHOW UAC PROMPT!) <-- true!
					//(Via context-menu "Repair"
					// command in Windows Explorer)
					//
					//UPGRADE (old):					Str			-			Str			Str			InstallFldrPath			-						S-1-5-18 = Local System = (Shows UAC prompt)
					//(IMPORTANT: Called from MSI of
					// previous/old version!)
					// and then
					//
					//UPGRADE (new):					-			-			-			-			InstallFldrPath			Yes						S-1-5-18 = Local System = (Shows UAC prompt)
					//(Called from new MSI version)

					if (!isStagePartOn(arrParts, IDX_PART.IDX_P_INSTALLED) && !isStagePartOn(arrParts, IDX_PART.IDX_P_REINSTALL) &&
						!isStagePartOn(arrParts, IDX_PART.IDX_P_UPGRADEPRODUCTCODE) && !isStagePartOn(arrParts, IDX_PART.IDX_P_REMOVE))
					{
						//Determine what stage we're on
						try
						{
							RegistryKey regKey = res.caller.IsSystem ? Registry.LocalMachine : Registry.CurrentUser;

							//Try to read the key first
							using (RegistryKey rKO = regKey.OpenSubKey(res.strRegistryKey, false))
							{
								bool bInstallFldrExisted;

								if (rKO != null)
								{
									//Read its value then
									bInstallFldrExisted = false;
									object objV = rKO.GetValue(kStrRegVal_InstallFolderExisted);
									if (objV != null)
									{
										//Convert it to integer
										int v;
										if (int.TryParse(objV.ToString(), out v))
											bInstallFldrExisted = v != 0;
									}

									res.stage = !bInstallFldrExisted ? MSI_STAGE.MS_INSTALL : MSI_STAGE.MS_UPGRADE_NEW;
								}
								else
								{
									//No key - then it's the first call to INSTALL stage

									//See if installation folder exists
									bInstallFldrExisted = Directory.Exists(res.strInstallFolder);

									//Set it in a temp registry
									using (RegistryKey rK = regKey.CreateSubKey(res.strRegistryKey))
									{
										rK.SetValue(kStrRegVal_InstallFolderExisted, bInstallFldrExisted ? 1 : 0);

										//Pick result for the stage
										res.stage = !bInstallFldrExisted ? MSI_STAGE.MS_INSTALL : MSI_STAGE.MS_UPGRADE_NEW;
									}
								}
							}

							if (bAfter)
							{
								//Delete registry value
								using (RegistryKey rKO = regKey.OpenSubKey(res.strRegistryKey, true))
								{
									if (rKO != null)
									{
										rKO.DeleteValue(kStrRegVal_InstallFolderExisted, false);

										if (rKO.ValueCount == 0)
										{
											regKey.DeleteSubKey(res.strRegistryKey);
										}
									}
								}
							}
						}
						catch (Exception ex)
						{
							//Something failed
							LogDiagnosticSpecError(1064, ex);
							res.stage = MSI_STAGE.MS_Unknown;
							return res;
						}
					}
					else if (/*isStagePartOn(arrParts, IDX_PART.IDX_P_INSTALLED) &&*/ !isStagePartOn(arrParts, IDX_PART.IDX_P_REINSTALL) &&
						!isStagePartOn(arrParts, IDX_PART.IDX_P_UPGRADEPRODUCTCODE) && isStagePartOn(arrParts, IDX_PART.IDX_P_REMOVE))
					{
						//Remove our temp reg value
						try
						{
							RegistryKey regKey = res.caller.IsSystem ? Registry.LocalMachine : Registry.CurrentUser;

							using (RegistryKey rKO = regKey.OpenSubKey(res.strRegistryKey, true))
							{
								if (rKO != null)
								{
									rKO.DeleteValue(kStrRegVal_InstallFolderExisted, false);

									if (rKO.ValueCount == 0)
									{
										regKey.DeleteSubKey(res.strRegistryKey);
									}
								}
							}
						}
						catch (Exception ex)
						{
							//Failed
							LogDiagnosticSpecError(1065, ex);
						}

						res.stage = MSI_STAGE.MS_UNINSTALL;
					}
					else if (isStagePartOn(arrParts, IDX_PART.IDX_P_INSTALLED) && !isStagePartOn(arrParts, IDX_PART.IDX_P_REINSTALL) &&
						!isStagePartOn(arrParts, IDX_PART.IDX_P_UPGRADEPRODUCTCODE) /*&& !isStagePartOn(arrParts, IDX_PART.IDX_P_REMOVE)*/)
					{
						res.stage = MSI_STAGE.MS_CHANGE;
					}
					else if (isStagePartOn(arrParts, IDX_PART.IDX_P_INSTALLED) && isStagePartOn(arrParts, IDX_PART.IDX_P_REINSTALL) &&
						!isStagePartOn(arrParts, IDX_PART.IDX_P_UPGRADEPRODUCTCODE) && !isStagePartOn(arrParts, IDX_PART.IDX_P_REMOVE))
					{
						res.stage = MSI_STAGE.MS_REPAIR;
					}
					else if (isStagePartOn(arrParts, IDX_PART.IDX_P_INSTALLED) && !isStagePartOn(arrParts, IDX_PART.IDX_P_REINSTALL) &&
						isStagePartOn(arrParts, IDX_PART.IDX_P_UPGRADEPRODUCTCODE) && isStagePartOn(arrParts, IDX_PART.IDX_P_REMOVE))
					{
						if (!bAfter)
						{
							try
							{
								//See if installation folder exists
								bool bInstallFldrExisted = Directory.Exists(res.strInstallFolder);

								RegistryKey regKey = res.caller.IsSystem ? Registry.LocalMachine : Registry.CurrentUser;

								//Set it in a temp registry
								using (RegistryKey rK = regKey.CreateSubKey(res.strRegistryKey))
								{
									rK.SetValue(kStrRegVal_InstallFolderExisted, bInstallFldrExisted ? 1 : 0);
								}
							}
							catch (Exception ex)
							{
								//Failed
								LogDiagnosticSpecError(1066, ex);
								res.stage = MSI_STAGE.MS_Unknown;
								return res;
							}
						}

						res.stage = MSI_STAGE.MS_UPGRADE_PREV;
					}


				}
				else
					LogDiagnosticSpecError(1063, "cnt=" + nCntPrts);
			}
			else
				LogDiagnosticSpecError(1062, "cnt=" + nCnt);

			return res;
		}

		private static bool isStagePartOn(string[] arrParts, IDX_PART idx)
		{
			return !string.IsNullOrEmpty(arrParts[(int)idx]);
		}

		private static string makeFolderEndWithSlash(string strPath, bool bAlwaysSlashTerminated)
		{
			strPath = strPath.TrimEnd(new char[] { '\\', '/' });
			if (bAlwaysSlashTerminated)
				strPath += Path.DirectorySeparatorChar;

			return strPath;
		}



		private static bool registerEventSource()
		{
			//Register event source for logging messages
			bool bRes = false;

			try
			{
				if (!EventLog.SourceExists(GlobalDefs.GlobDefs.gkstrEventSrcName))
				{
					EventLog.CreateEventSource(GlobalDefs.GlobDefs.gkstrEventSrcName, "Application");

					bRes = true;
				}
			}
			catch
			{
				//Failed - no need to log it as we don't have the event source :(
				bRes = false;
			}

			return bRes;
		}

		private static bool unregisterEventSource()
		{
			//Remove event source
			bool bRes = false;

			try
			{
				if (!EventLog.SourceExists(GlobalDefs.GlobDefs.gkstrEventSrcName))
				{
					EventLog.DeleteEventSource(GlobalDefs.GlobDefs.gkstrEventSrcName);

					bRes = true;
				}
			}
			catch
			{
				//Failed
				bRes = false;
			}

			return bRes;
		}

		public static bool isThis64bitProc()
		{
			//RETURN: = true if we're running as a 64-bit process
			return IntPtr.Size == 8;
		}



		public static bool LogDiagnosticEvent(int nSpecErr, string strMsg, EventLogEntryType type)
		{
			//Places a diagnostic message into an event log
			//INFO: This function will add current time & version of the app
			//'nSpecErr' = special unique error ID in the source code. To generate it, use SeqIDGen: https://dennisbabkin.com/seqidgen
			//'strMsg' = additional text message to log to describe an error
			//'type' = type of an error message
			//RETURN:
			//		= true if logged without any errors
			bool bRes = false;

			try
			{
				using (EventLog eventLog = new EventLog("Application"))
				{
					DateTime dtNow = DateTime.Now;

					eventLog.Source = GlobalDefs.GlobDefs.gkstrEventSrcName;

					eventLog.WriteEntry(dtNow.Year.ToString("D4") + "-" + dtNow.Month.ToString("D2") + "-" + dtNow.Day.ToString("D2") + " " +
						dtNow.Hour.ToString("D2") + ":" + dtNow.Minute.ToString("D2") + ":" + dtNow.Second.ToString("D2") + "." + dtNow.Millisecond.ToString("D3") +
						": [Installer - " + GlobalDefs.GlobDefs.gkstrAppNameFull + " v." + gAppVersion.getVersion() + "]" +
						"[" + (isThis64bitProc() ? "x64" : "x86") + ">" + Environment.UserName + "]" +
						"[" + nSpecErr + "]" +
						(!string.IsNullOrEmpty(strMsg) ? " " : "") + strMsg,
						type, 1, 0);

					bRes = true;
				}
			}
			catch (Exception)
			{
				//Failed
				bRes = false;
			}

			return bRes;
		}

		public static bool LogDiagnosticSpecError(int nSpecErr)
		{
			//Places a diagnostic error message into an event log
			//INFO: This function will add current time & version of the app
			//'nSpecErr' = special unique error ID in the source code. To generate it, use SeqIDGen: https://dennisbabkin.com/seqidgen
			//RETURN:
			//		= true if logged without any errors
			return LogDiagnosticEvent(nSpecErr, "", EventLogEntryType.Error);
		}

		public static bool LogDiagnosticSpecError(int nSpecErr, Exception ex)
		{
			//Places a diagnostic error message into an event log
			//INFO: This function will add current time & version of the app
			//'nSpecErr' = special unique error ID in the source code. To generate it, use SeqIDGen: https://dennisbabkin.com/seqidgen
			//'ex' = exception with error description
			//RETURN:
			//		= true if logged without any errors
			return LogDiagnosticEvent(nSpecErr, ex.ToString(), EventLogEntryType.Error);
		}

		public static bool LogDiagnosticSpecError(int nSpecErr, Session session)
		{
			//Places a diagnostic error message into an event log
			//INFO: This function will add current time & version of the app
			//'nSpecErr' = special unique error ID in the source code. To generate it, use SeqIDGen: https://dennisbabkin.com/seqidgen
			//'session' = installation session
			//RETURN:
			//		= true if logged without any errors

			string str = "";
			int nCnt = session.CustomActionData.Count;
			str += "MSI_SESSION: [" + nCnt.ToString() + "]";
			if (nCnt != 0)
			{
				str += ": " + session.CustomActionData.ToString();
			}

			return LogDiagnosticEvent(nSpecErr, str, EventLogEntryType.Error);
		}

		

		public static bool LogDiagnosticSpecError(int nSpecErr, string strMsg)
		{
			//Places a diagnostic error message into an event log
			//INFO: This function will add current time & version of the app
			//'nSpecErr' = special unique error ID in the source code. To generate it, use SeqIDGen: https://dennisbabkin.com/seqidgen
			//'strMsg' = additional text message to log to describe an error
			//RETURN:
			//		= true if logged without any errors
			return LogDiagnosticEvent(nSpecErr, strMsg, EventLogEntryType.Error);
		}

		public static bool LogDiagnosticSpecWarning(int nSpecErr)
		{
			//Places a diagnostic warning message into an event log
			//INFO: This function will add current time & version of the app
			//'nSpecErr' = special unique error ID in the source code. To generate it, use SeqIDGen: https://dennisbabkin.com/seqidgen
			//RETURN:
			//		= true if logged without any errors
			return LogDiagnosticEvent(nSpecErr, "", EventLogEntryType.Warning);
		}

		public static bool LogDiagnosticSpecWarning(int nSpecErr, string strMsg)
		{
			//Places a diagnostic warning message into an event log
			//INFO: This function will add current time & version of the app
			//'nSpecErr' = special unique error ID in the source code. To generate it, use SeqIDGen: https://dennisbabkin.com/seqidgen
			//'strMsg' = additional text message to log to describe a warning
			//RETURN:
			//		= true if logged without any errors
			return LogDiagnosticEvent(nSpecErr, strMsg, EventLogEntryType.Warning);
		}


	
	
	
	
	}
}
