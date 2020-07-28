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
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

using System.Diagnostics;
using System.Windows.Forms;
using System.Xml;
using System.Reflection;
using System.Media;
using System.Configuration;
using System.IO;
using Microsoft.Win32;
using System.Runtime.InteropServices;



namespace OutlookHeaders
{
    public partial class ThisAddIn
    {
        private const string kstrMainMenuTag = "{5965471F-28B3-491a-AAA1-5626FDA5FDD7}";

		public const string kstrDataFileVer = "1.0.0";		//XML export data file version produced by this version of the app (NOTE that this is NOT the same as the app version!)
		public const string kstrAppDownloadURL = "https://dennisbabkin.com/olh";
		public const string kstrAppManualURL = "https://dennisbabkin.com/php/docs.php?what=olh&ver=#VER#&t=#TYPE#";
		public const string kstrAppUpdatesURL = "https://dennisbabkin.com/php/update.php?name=olh&ver=#VER#&t=#TYPE#";
		public const string kstrAppBugReportURL = "https://dennisbabkin.com/sfb/?what=bug&name=#NAME#&ver=#VER#&msg=#MSG#";

		public static CurrentVersion gAppVersion = new CurrentVersion();			//Will be added later from the assembly file

		public StartupFlags gStartupFlgs = StartupFlags.STUP_FLG_None;				//Start-up flags - can be provided on the input during installation from an MSI installer


        //TROUBLESHOOTING:
        //      https://stackoverflow.com/questions/4668777/how-to-troubleshoot-a-vsto-addin-that-does-not-load


		public class StrNameVal
		{
			public string strName;
			public string strVal;
		}

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



        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
			//Check if we need to upgrade config file and do so if needed
			UpgradeConfigIfNeeded();
			UpgradeConfigAppVersionIfNeeded();


            //Add our UI menu
            AddMenuBar();

			//Set event for the "Send" method to catch when emails are being sent out
            this.Application.ItemSend += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ItemSendEventHandler(onItemSend);

			//Set event handler for additional events
			this.Application.ItemLoad += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ItemLoadEventHandler(onItemLoad);

        }


		HashSet<Outlook.ItemEvents_10_Event> garrObj_itemEvts = new HashSet<Microsoft.Office.Interop.Outlook.ItemEvents_10_Event>();

		private void onItemLoad(object Item)
		{
			//This processes multiple events
			try
			{
				//Don't do it multiple times
				Outlook.ItemEvents_10_Event itemEvt = (Outlook.ItemEvents_10_Event)Item;
				if (itemEvt != null)
				{
					//Handlers for Reply, Reply-All and Forward
					itemEvt.Reply += new Microsoft.Office.Interop.Outlook.ItemEvents_10_ReplyEventHandler(on_Reply);
					itemEvt.ReplyAll += new Microsoft.Office.Interop.Outlook.ItemEvents_10_ReplyAllEventHandler(on_ReplyAll);
					itemEvt.Forward += new Microsoft.Office.Interop.Outlook.ItemEvents_10_ForwardEventHandler(on_Forward);

					itemEvt.Unload += new Microsoft.Office.Interop.Outlook.ItemEvents_10_UnloadEventHandler(delegate()
						{
							try
							{
								//Remove this event from our hashset
								bool bR = garrObj_itemEvts.Remove(itemEvt);

								//LogDiagnosticSpecWarning(1, "on_ItemUnload=" + bR.ToString());
							}
							catch (Exception ex)
							{
								//Failed
								LogDiagnosticSpecError(1154, ex);
							}
						});

					//Remember in the global variable
					garrObj_itemEvts.Add(itemEvt);

					//LogDiagnosticSpecWarning(1, "onItemLoad cnt=" + garrObj_itemEvts.Count);
				}
				else
					LogDiagnosticSpecError(1153);
			}
			catch (Exception ex)
			{
				//Failed
				LogDiagnosticSpecError(1152, ex);
			}
		}




		private void on_Reply(object Item, ref bool cancel)
		{
			//Called when user loads Reply window
			adjustMessageFormatting(Item, EmlMsgType.EMT_Reply);
		}

		private void on_ReplyAll(object Item, ref bool cancel)
		{
			//Called when user loads Reply-All window
			adjustMessageFormatting(Item, EmlMsgType.EMT_ReplyAll);
		}

		private void on_Forward(object Item, ref bool cancel)
		{
			//Called when user loads Forward window
			adjustMessageFormatting(Item, EmlMsgType.EMT_Forward);
		}

        private void onItemSend(object Item, ref bool Cancel)
        {
			//This method is called when an email is being sent from Outlook
			//'Item' = email object that is being sent
            Outlook.MailItem objMailItem = (Outlook.MailItem)Item;

			bool bLog = false;
			bool bEnabled = false;

			string strEmailFrom = "";
			string strEmailTo = "";
			string strSubject = "";
			Outlook.OlBodyFormat msgOrigFmt = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatUnspecified;

			List<StrNameVal> arrAllAddedHdrs = new List<StrNameVal>();		//List of all name=value pairs that were added to this email headers

			FormDefineHeaders.OutboundEmailInfo emlFmt = new FormDefineHeaders.OutboundEmailInfo();
			ProcessedBitmask procBtmk = ProcessedBitmask.PBMSK_None;

			string strSingleEmailRecip = "";		//Only first email recipient (used for logging purposes)

			string strErrDesc = "";
			int nSpecErr = 0;

			try
			{
				//Get email details
				strEmailFrom = objMailItem.SendUsingAccount.SmtpAddress;
				strEmailTo = objMailItem.To;
				strSubject = objMailItem.Subject;
				msgOrigFmt = objMailItem.BodyFormat;

				//Read config
				bLog = Properties.Settings.Default.OutlookHdrsLogWhenSending;
				bEnabled = Properties.Settings.Default.OutlookHdrsEnabled;

				if (objMailItem.Recipients != null)
				{
					//Add recipients email addresses
					string strRecipEmls = "";
					int nCntRcpt = objMailItem.Recipients.Count;
					for (int r = 1; r <= nCntRcpt; r++)				//Microsoft is "special" - they count from index 1!
					{
						string strEml = objMailItem.Recipients[r].Address;
						if (!string.IsNullOrEmpty(strEml))
						{
							if (!string.IsNullOrEmpty(strRecipEmls))
								strRecipEmls += "; ";

							strRecipEmls += strEml;

							if(string.IsNullOrEmpty(strSingleEmailRecip))
								strSingleEmailRecip = strEml;
						}
					}

					if (!string.IsNullOrEmpty(strRecipEmls))
					{
						strEmailTo += " <" + strRecipEmls + ">";
					}
				}

				//Only if add-in is enabled
				if (bEnabled)
				{
					//Get headers to process
					XmlDocument xmlHdrs = readHeaders(out strErrDesc);
					if (xmlHdrs != null)
					{
						XmlElement xmlRoot = xmlHdrs.DocumentElement;
						if (xmlRoot == null)
							new Exception("[1046]");

						//Find the <ALL> node first
						XmlNode xNodeAll = xmlRoot.SelectSingleNode(FormDefineHeaders.kgStrNodeNm_Account + "[@" + FormDefineHeaders.kgStrNodeAttNm_All + "]");
						if (xNodeAll == null)
							new Exception("[1047]");	//We must have the ALL node!

						//Process headers for the <ALL> node
						processNodeHeadersForSending(objMailItem, xNodeAll, "<ALL>", strSingleEmailRecip, ref arrAllAddedHdrs, ref emlFmt, ref procBtmk);


						XmlAttribute xAtt;

						//Then go through the other nodes and match them to our account
						foreach (XmlNode xNode in xmlRoot.ChildNodes)
						{
							//Skip <ALL> node
							if (xNode != xNodeAll)
							{
								xAtt = xNode.Attributes[FormDefineHeaders.kgStrNodeAttNm_Name];
								string strName = (xAtt != null) ? xAtt.Value : null;

								if (!string.IsNullOrEmpty(strName) &&
									string.Compare(strName, strEmailFrom, true) == 0)
								{
									//Match -- process these headers
									processNodeHeadersForSending(objMailItem, xNode, strName, strSingleEmailRecip, ref arrAllAddedHdrs, ref emlFmt, ref procBtmk);

									break;
								}
							}
						}



					}
					else
					{
						//Failed to read headers - rethrow
						new Exception("[1045] " + strErrDesc);
					}
				}

				//Do we need to log it?
				if (bLog)
				{
					//Log this event
					nSpecErr = 1044;
				}
			}
			catch (Exception ex)
			{
				//Exception
				nSpecErr = -1043;
				strErrDesc = ex.ToString();
			}


			//See if we need to log it?
			if (nSpecErr != 0)
			{
				//Turn headers added into a string
				string strAllHdrs = "";
				foreach (StrNameVal snv in arrAllAddedHdrs)
				{
					strAllHdrs += snv.strName + ": " + snv.strVal + "\n";
				}

				//Log it (negative spec-error means errors)
				ThisAddIn.LogDiagnosticEvent(Math.Abs(nSpecErr),
					"FROM: " + strEmailFrom + "\n" +
					"TO: " + strEmailTo + "\n" +
					"SUBJ: " + strSubject + "\n" +
					"FMT: " + msgOrigFmt.ToString() + "\n" +
					"MAIL_CLIENT: " + ((procBtmk & ProcessedBitmask.PBMSK_OverwriteMailClient) != 0 ? emlFmt.strMailClientName : "<NotUsed>") + "\n" +
					"RCVD_HDR: " + ((procBtmk & ProcessedBitmask.PBMSK_SuppressReceivedHeader) != 0 ? parseReceivedHeaderValue(emlFmt.strRcvdHeader, strSingleEmailRecip) : "<NotUsed>") + "\n" +
					"HDRS_APPLIED:\n" + 
					(bEnabled ? strAllHdrs : "<Add-in Disabled>\n" +
					(!string.IsNullOrEmpty(strErrDesc) ? "\nERROR: " + strErrDesc : "")
					)
					,
					nSpecErr > 0 ? EventLogEntryType.Information : EventLogEntryType.Error);
			}
		}


		private enum ProcessedBitmask
		{
			PBMSK_None = 0,
			PBMSK_OverwriteMailClient = 0x1,
			PBMSK_SuppressReceivedHeader = 0x2,
		}


		private static bool processNodeHeadersForSending(Outlook.MailItem objMailItem, XmlNode xNode, string strAccoutName, string strEmailTo, ref List<StrNameVal> arrAllAddedHdrs, ref FormDefineHeaders.OutboundEmailInfo emlFmt, ref ProcessedBitmask procBtmsk)
		{
			//Add headers for the account
			//'objMailItem' = mail object that is being sent
			//'xNode' = XML node with config data with headers for this account
			//'strAccoutName' = account name (used for debugging purposes only)
			//'strEmailTo' = who this message is set to - email address (only one recipient is enough)
			//'arrAllAddedHdrs' = array that receives name=value pairs of all headers that were added
			//'procBtmsk' = bits of what has been done already
			//RETURN:
			//		= true if all added OK
			bool bRes = true;

			//See if we need to change format of the email
			processNodeHeadersForFormatChange(xNode, ref emlFmt);

			//Go through all the headers in the node
			foreach (XmlNode xNHdr in xNode)
			{
				FormDefineHeaders.ItemInfo iiHdr = FormDefineHeaders.getHeaderInfoFromNode(xNHdr);
				if (iiHdr != null)
				{
					//Only if enabled by the user
					if (iiHdr.bEnabled)
					{
						//Add it
						try
						{
							string strEncName = iiHdr.strName, strEncVal = iiHdr.strVal;
							//HeaderNameValueEncode(iiHdr.strName, iiHdr.strVal, out strEncName, out strEncVal);

							//SOURCE:
							//	https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxprops/cc9d955b-1492-47de-9dce-5bdea80a3323

							//We need special handling here... why? Because Microsoft are special
							string PS_INTERNET_HEADERS = "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/" +
								strEncName;

							objMailItem.PropertyAccessor.SetProperty(PS_INTERNET_HEADERS, strEncVal);

							//Add this to our log
							arrAllAddedHdrs.Add(new StrNameVal() { strName = strEncName, strVal = strEncVal });


							//Do we need to overwrite the mail client name
							if (emlFmt.useOverwriteMailClient == FormDefineHeaders.TriState.TS_True &&
								(procBtmsk & ProcessedBitmask.PBMSK_OverwriteMailClient) == 0)
							{
								//Set mail client name
								objMailItem.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/X-Mailer",
									emlFmt.strMailClientName);

								procBtmsk |= ProcessedBitmask.PBMSK_OverwriteMailClient;
							}

							//Do we need to suppress Received header
							if (emlFmt.useSuppressRcvdHeader == FormDefineHeaders.TriState.TS_True &&
								(procBtmsk & ProcessedBitmask.PBMSK_SuppressReceivedHeader) == 0)
							{
								objMailItem.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/Received",
									parseReceivedHeaderValue(emlFmt.strRcvdHeader, strEmailTo));

								procBtmsk |= ProcessedBitmask.PBMSK_SuppressReceivedHeader;
							}						
						}
						catch (Exception ex)
						{
							//Failed to add this header
							ThisAddIn.LogDiagnosticSpecError(1048, "ACCT: " + strAccoutName + ", NM: " + iiHdr.strName + ", VAL: " + iiHdr.strVal + " > " + ex);

							bRes = false;
						}
					}
				}
				else
				{
					//Failed to read this header
					ThisAddIn.LogDiagnosticSpecError(1049, "ACCT: " + strAccoutName);

					bRes = false;
				}
			}

			return bRes;
		}

		public static string parseReceivedHeaderValue(string strVal, string strDestEmailAddr)
		{
			//'strVal' = template string to parse
			//'strDestEmailAddr' = destination email address
			//RETURN:
			//		= Resulting string

			try
			{
				// /from [192.168.1.124] (host-name.com [<ip>]) by pv50p00im-tydg10021701.me.com (Postfix) with ESMTPSA id 013098402CF for <dest-addr@gmail.com>; Sun, 12 Jul 2020 00:09:54 +0000 (UTC)
				strVal = strVal.Replace("#DEST_EMAIL_ADDR#", strDestEmailAddr);

				//Replace UTC date: Sun, 12 Jul 2020 00:09:54 +0000 (UTC)
				DateTime dtNowUtc = DateTime.UtcNow;
				string[] strDOWs = { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };
				string[] strMonths = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };

				string strDtm = strDOWs[(int)dtNowUtc.DayOfWeek] + ", " + dtNowUtc.Day.ToString("D2") + " " + strMonths[dtNowUtc.Month - 1] + " " + dtNowUtc.Year.ToString("D4") + " " +
					dtNowUtc.Hour.ToString("D2") + ":" + dtNowUtc.Minute.ToString("D2") + ":" + dtNowUtc.Second.ToString("D2") + " +0000 (UTC)";

				strVal = strVal.Replace("#DATE_UTC#", strDtm);
			}
			catch (Exception ex)
			{
				LogDiagnosticSpecError(1151, ex);
			}

			return strVal;
		}


		public static void HeaderNameValueEncode(string headerName, string headerValue, out string encodedHeaderName, out string encodedHeaderValue)
		{
			//Have to roll our own encoder, as we don't have .NET 4.0 here

			//SOURCE:
			//		https://stackoverflow.com/questions/2769080/how-to-encode-custom-http-headers-in-c-sharp

			if (String.IsNullOrEmpty(headerName))
				encodedHeaderName = "";
			else
				encodedHeaderName = encodeHeaderValueStr(headerName, true);

			if (String.IsNullOrEmpty(headerValue))
				encodedHeaderValue = "";
			else
				encodedHeaderValue = encodeHeaderValueStr(headerValue, false);
		}

		private static string encodeHeaderValueStr(string strInput, bool bName)
		{
			//'strInput' = string to encode
			//'bName' = true if 'strInput' is mail header name, false if it's mail header value (or field body)
			//RETURN:
			//		= Escaped string
			string strRes = "";

			int nLn = strInput.Length;
			for (int i = 0; i < nLn; i++)
			{
				char ch = strInput[i];

				bool bEncode;
				if (bName)
				{
					bEncode = ch < 33 || ch > 126 || ch == '%' || ch == '/' || ch == '\\';
				}
				else
				{
					bEncode = (ch < 32 && ch != '\t') || ch > 126 || ch == '%' || ch == '/' || ch == '\\';
				}

				if (bEncode)
				{
					strRes += String.Format("%{0:x2}", (int)ch);
				}
				else
					strRes += ch;
			}

			return strRes;
		}


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
			////Remove our menu
			//RemoveMenuBar();

        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion


		//IMPORTANT:
		//		We need these to be on the global scale so that the GarbageCollector doesn't remove them
		//		in which case, our menus will stop working!
		//		Garbage Collecter in this case evidently decides not to keep references to objects and deletes them.... hah?!
		Office.CommandBar g_gbjCmdBar = null;
		Office.CommandBarPopup g_objMenuBar = null;
		Office.CommandBarButton g_objBtnEnabled = null;
		Office.CommandBarButton g_objBtnDefineHdrs = null;


        private bool AddMenuBar()
        {
            //Add our menu bar
            bool bRes = false;

            //First remove existing
            RemoveMenuBar();

            try
            {
				g_gbjCmdBar = this.Application.ActiveExplorer().CommandBars.ActiveMenuBar;

                //Add new
				//IMPORTANT: Make sure to set Temporary property to 'true' so that control is removed automatically when we uninstall it!
				//           Another gotcha that took half-a-day to discover .... :(
				g_objMenuBar = (Office.CommandBarPopup)
					g_gbjCmdBar.Controls.Add(Office.MsoControlType.msoControlPopup, missing, missing, missing, true);
				if (g_objMenuBar != null)
                {
					//Read config property value
					bool bEnabled = Properties.Settings.Default.OutlookHdrsEnabled;

					g_objMenuBar.Caption = getMainMenuName(bEnabled);
					g_objMenuBar.Tag = kstrMainMenuTag;

                    //Add "enabled" checkbox
					g_objBtnEnabled = (Office.CommandBarButton)
						g_objMenuBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, g_objMenuBar.Controls.Count + 1, true);

					g_objBtnEnabled.Caption = "&Enabled";
					g_objBtnEnabled.Style = Microsoft.Office.Core.MsoButtonStyle.msoButtonCaption;
					g_objBtnEnabled.Tag = "tagEnabled";
					g_objBtnEnabled.State = bEnabled ? Microsoft.Office.Core.MsoButtonState.msoButtonDown : Microsoft.Office.Core.MsoButtonState.msoButtonUp;
					g_objBtnEnabled.Click += new Office._CommandBarButtonEvents_ClickEventHandler(onButton_Click_Enabled);

                    //Add button to open config window
					g_objBtnDefineHdrs = (Office.CommandBarButton)
						g_objMenuBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, g_objMenuBar.Controls.Count + 1, true);

					g_objBtnDefineHdrs.Caption = "Customize Outbound &Mail...";
					g_objBtnDefineHdrs.Style = Microsoft.Office.Core.MsoButtonStyle.msoButtonCaption;
					g_objBtnDefineHdrs.Tag = "tagDefineHdrs";
					g_objBtnDefineHdrs.Click += new Office._CommandBarButtonEvents_ClickEventHandler(onButton_Click_DefineHeaders);
					g_objBtnDefineHdrs.BeginGroup = true;       //Use as a separator

                }
				else
					ThisAddIn.LogDiagnosticSpecError(1001);
			}
            catch (Exception ex)
            {
				ThisAddIn.LogDiagnosticSpecError(1000, ex);
				MessageBox.Show("ERROR: Failed to add new menu bar" + ex.ToString(), GlobalDefs.GlobDefs.gkstrAppNameFull, MessageBoxButtons.OK, MessageBoxIcon.Error);
                bRes = false;
            }

            return bRes;
        }

        private static string getMainMenuName(bool bEnabled)
        {
			//RETURN:
			//		= Name of the main menu for the add-in, that is added to the Outllok
            string strName = "&Mail Headers";
            if (!bEnabled)
                strName += " (Disabled)";

            return strName;
        }

		public static XmlDocument readHeaders(out string strOutErrDesc)
		{ 
			//Read all headers from persistent storage
			//RETURN:
			//		= XML object with data
			//		= null if error (check 'strOutErrDesc' for details)
			strOutErrDesc = "";
			XmlDocument resDoc = null;

			try
			{
				resDoc = new XmlDocument();

				//Get string with XML
				string strXml = Properties.Settings.Default.OutlookHdrsMailHdrs;
				if (!string.IsNullOrEmpty(strXml))
				{
					resDoc.LoadXml(strXml);
				}

			}
			catch (Exception ex)
			{
				ThisAddIn.LogDiagnosticSpecError(1002, ex);
				strOutErrDesc = "ERROR: " + ex.ToString();
				resDoc = null;
			}

			return resDoc;
		}

		public static bool saveHeaders(string strXml, out string strOutErrDesc)
		{
			//Save XML headers data into persistent storage
			//'strXml' = XML data to save
			//RETURN:
			//		= true if success
			//		= null if error (check 'strOutErrDesc' for details)
			strOutErrDesc = "";
			bool bRes = false;

			try
			{
				//Set it in the config properties
				Properties.Settings.Default.OutlookHdrsMailHdrs = strXml;
				Properties.Settings.Default.Save();

				//Done
				bRes = true;
			}
			catch (Exception ex)
			{
				//Exception
				ThisAddIn.LogDiagnosticSpecError(1003, ex);
				strOutErrDesc = "ERROR: " + ex.ToString();
				bRes = false;
			}

			return bRes;
		}


		private bool RemoveMenuBar()
        {
            //Remove our menu bar
            bool bRes = false;

            try
            {
				if(this.Application != null)
				{
					Microsoft.Office.Interop.Outlook.Explorer explr = this.Application.ActiveExplorer();
					if (explr != null)
					{
						Microsoft.Office.Core.CommandBars cmdBars = explr.CommandBars;
						if (cmdBars != null)
						{
							Microsoft.Office.Core.CommandBar actvBar = cmdBars.ActiveMenuBar;
							if (actvBar != null)
							{
								Office.CommandBarPopup objMenuItm = (Office.CommandBarPopup)
									actvBar.FindControl(Office.MsoControlType.msoControlPopup, missing, kstrMainMenuTag, true, true);
								if (objMenuItm != null)
								{
									objMenuItm.Delete(true);

									bRes = true;
								}
							}
						}
					}
				}
            }
            catch (Exception ex)
            {
				ThisAddIn.LogDiagnosticSpecError(1004, ex);
				MessageBox.Show("ERROR: Failed to remove menu bar: " + ex.ToString(), GlobalDefs.GlobDefs.gkstrAppNameFull, MessageBoxButtons.OK, MessageBoxIcon.Error);
                bRes = false;
            }

            return bRes;
        }

		private void onButton_Click_Enabled(Office.CommandBarButton ctrl, ref bool cancel)
		{
			//Enable/Disable state changed
			bool bEnabled = ctrl.State == Microsoft.Office.Core.MsoButtonState.msoButtonUp;

			try
			{
				//Set it in the config properties
				Properties.Settings.Default.OutlookHdrsEnabled = bEnabled;
				Properties.Settings.Default.Save();

				//Set new state
				ctrl.State = bEnabled ? Microsoft.Office.Core.MsoButtonState.msoButtonDown : Microsoft.Office.Core.MsoButtonState.msoButtonUp;

				//And rename main menu
				Office.CommandBarPopup objMenuItm = (Office.CommandBarPopup)
					this.Application.ActiveExplorer().CommandBars.ActiveMenuBar.FindControl(Office.MsoControlType.msoControlPopup, missing, kstrMainMenuTag, true, true);
				if (objMenuItm != null)
				{
					objMenuItm.Caption = getMainMenuName(bEnabled);
				}
				else
					ThisAddIn.LogDiagnosticSpecError(1006);
			}
			catch (Exception ex)
			{
				//Failed
				ThisAddIn.LogDiagnosticSpecError(1005, ex);
				MessageBox.Show("ERROR: Failed to save setting.\n\n" + ex.ToString(), GlobalDefs.GlobDefs.gkstrAppNameFull, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

        private void onButton_Click_DefineHeaders(Office.CommandBarButton ctrl, ref bool cancel)
        {
			//Show main dialog
            FormDefineHeaders fdh = new FormDefineHeaders();
            fdh._this = this;
			DialogResult resDlg = fdh.ShowDialog();
			if (resDlg != DialogResult.OK &&
				resDlg != DialogResult.Cancel)
			{
				//Some kind of error
				ThisAddIn.LogDiagnosticSpecError(1159, "res=" + resDlg.ToString());
				SystemSounds.Exclamation.Play();
			}
        }

		public static string getThisProjectName()
		{
			//RETURN:
			//		= String name of this C# project, or
			//		= "" if error
			try
			{
				return System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
			}
			catch (Exception ex)
			{
				//Failed
				ThisAddIn.LogDiagnosticSpecError(1054, ex);

				return "";
			}
		}


		[DllImport("kernel32.dll", SetLastError = true)]
		private static extern bool IsWow64Process(
			[In] IntPtr hProcess,
			[Out] out bool wow64Process
		);

		public static bool is32bitProcessOn64bitOS()
		{
			//RETURN:
			//		= true if we're running on a 64-bit OS as a 32-bit process

			//Check that we're at least on Windows XP SP2
			if (Environment.OSVersion.Version.Major >= 6 ||
				(Environment.OSVersion.Version.Major == 5 && Environment.OSVersion.Version.Minor >= 1))
			{
				using (Process proc = Process.GetCurrentProcess())
				{
					bool bWow64;
					if (IsWow64Process(proc.Handle, out bWow64))
						return bWow64;
				}
			}

			return false;
		}

		public static bool isThis64bitProc()
		{
			//RETURN: = true if we're running as a 64-bit process
			return IntPtr.Size == 8;
		}


		public class HKLMInstallerData
		{
			public string strImportFile = null;			//Not-null if data was read
			public Nullable<uint> flags = null;			//Not-null if data was read
		}

		[DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
		public static extern int RegOpenKeyEx(
		  UIntPtr hKey,
		  string subKey,
		  int ulOptions,
		  int samDesired,
		  out UIntPtr hkResult);

		[DllImport("advapi32.dll", SetLastError = true)]
		public static extern int RegCloseKey(
			UIntPtr hKey);

		[DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
		static extern uint RegQueryValueEx(
			UIntPtr hKey,
			string lpValueName,
			int lpReserved,
			ref RegistryValueKind lpType,
			IntPtr lpData,
			ref int lpcbData);


		public static HKLMInstallerData getInstallerHKLMData()
		{
			//INFO:
			//     Have to write this one using native WinAPIs since .NET 3.5 does not support
			//     64-bit System Registry Redirection features :(
			//
			//RETURN:
			//		= Values read from HKLM hive of the registry (that are placed there by an Installer)
			//		= null if none exist
			HKLMInstallerData resData = null;

			//Is it a 64-bit process?
			bool b64bit = isThis64bitProc();

			UIntPtr hKey;
			if (RegOpenKeyEx(new UIntPtr(0x80000002u),		//HKEY_LOCAL_MACHINE                  (( HKEY ) (ULONG_PTR)((LONG)0x80000002) )
				GlobalDefs.GlobDefs.gkstrAppRegistryKey,
				0,
				0x20019 |						//KEY_READ
				(b64bit ? 0x0200 : 0),			//KEY_WOW64_32KEY
				out hKey) == 0)
			{
				resData = new HKLMInstallerData();

				//Read flags
				RegistryValueKind type = RegistryValueKind.Unknown;
				int ncbSz = 4;
				IntPtr pResult = Marshal.AllocHGlobal(ncbSz);
				if (RegQueryValueEx(hKey, GlobalDefs.GlobDefs.gkstrAppRegVal_1stRun_StartupFlags, 0, ref type, pResult, ref ncbSz) == 0 &&
					type == RegistryValueKind.DWord &&
					ncbSz == 4)
				{
					//Got something
					resData.flags = new uint();
					resData.flags = (uint)Marshal.ReadInt32(pResult);
				}

				//Read path, but first get its size
				type = RegistryValueKind.Unknown;
				ncbSz = 0;
				RegQueryValueEx(hKey, GlobalDefs.GlobDefs.gkstrAppRegVal_1stRun_ImportDataFile, 0, ref type, IntPtr.Zero, ref ncbSz);
				if (ncbSz != 0 &&
					(type == RegistryValueKind.String || type == RegistryValueKind.ExpandString))
				{
					pResult = Marshal.AllocHGlobal(ncbSz);

					int ncbSz2 = ncbSz;
					if (RegQueryValueEx(hKey, GlobalDefs.GlobDefs.gkstrAppRegVal_1stRun_ImportDataFile, 0, ref type, pResult, ref ncbSz2) == 0 &&
						ncbSz2 == ncbSz)
					{
						if (type == RegistryValueKind.String || type == RegistryValueKind.ExpandString)
						{
							//Got something
							resData.strImportFile = Marshal.PtrToStringAuto(pResult);
						}
					}
				}

				//Free mem
				if (pResult != IntPtr.Zero)
				{
					Marshal.FreeHGlobal(pResult);
					pResult = IntPtr.Zero;
				}

				//Close the key
				RegCloseKey(hKey);
			}

			return resData;
		}


		public enum StartupFlags : uint
		{
			STUP_FLG_None = 0x0,
			STUP_FLG_START_DISABLED = 0x1,				//To start this add-in disabled
			STUP_FLG_LOG_WHEN_SENDING_EMAILS = 0x2,		//To log debugging info into diagnostic event log when emails are sent out
		}

		private bool GetInitialImportData(out string strImportDataFile, out StartupFlags startupFlags)
		{
			//Check if import data is available (that can be placed there by our MSI installer)
			//'strImportDataFile' = received XML data file to import (it will be verified)
			//'startupFlags' = received start-up flags
			//RETURN:
			//		= true if (some) imported data was found & validated
			bool bRes = false;

			strImportDataFile = "";
			startupFlags = StartupFlags.STUP_FLG_None;

			try
			{
				string strErrDesc;

				//Read data placed into the HKLM hive by the intaller
				HKLMInstallerData hklmInstDta = getInstallerHKLMData();
				if (hklmInstDta != null)
				{
					//Now refer to the HKCU keys
					using (RegistryKey regCU = Registry.CurrentUser.CreateSubKey(GlobalDefs.GlobDefs.gkstrAppRegistryKey))
					{
						if (hklmInstDta.strImportFile != null)
						{
							//If the HKCU value exists, then don't process it
							//INFO: We need it to prevent repeated processing of this setting when add-in restarts.
							if (regCU.GetValue(GlobalDefs.GlobDefs.gkstrAppRegVal_1stRun_ImportDataFile) != null)
							{
								hklmInstDta.strImportFile = null;
							}
						}

						if (hklmInstDta.flags != null)
						{
							//If the HKCU value exists, then don't process it
							//INFO: We need it to prevent repeated processing of this setting when add-in restarts.
							if (regCU.GetValue(GlobalDefs.GlobDefs.gkstrAppRegVal_1stRun_StartupFlags) != null)
							{
								hklmInstDta.flags = null;
							}
						}


						//See if we have our values
						if (hklmInstDta.strImportFile != null)
						{
							//Set key not to process it again
							regCU.SetValue(GlobalDefs.GlobDefs.gkstrAppRegVal_1stRun_ImportDataFile, "", RegistryValueKind.ExpandString);

							string strImportFile = hklmInstDta.strImportFile;
							strImportFile = strImportFile.Trim();

							if (File.Exists(strImportFile))
							{
								//Validate it
								string strXML = System.IO.File.ReadAllText(strImportFile);

								FormDefineHeaders.ImportedDoc impDoc = FormDefineHeaders.verifyXmlDataForImporting(strXML, out strErrDesc);
								if (impDoc != null)
								{
									//Use can use this file
									strImportDataFile = strImportFile;

									//Mark it that we've got something
									bRes = true;
								}
								else
								{
									//Failed to validate
									LogDiagnosticSpecError(1093, "Import file error: file=\"" + strImportFile + "\" > " + strErrDesc);
								}
							}
							else
							{
								//Bad file path
								LogDiagnosticSpecError(1094, "Import file couldn't be opened: file=\"" + strImportFile + "\"");
							}
						}

						//And flags?
						if (hklmInstDta.flags != null)
						{
							//Set special key to signify for us not to process it again
							uint uiFlags = hklmInstDta.flags.Value;
							regCU.SetValue(GlobalDefs.GlobDefs.gkstrAppRegVal_1stRun_StartupFlags, uiFlags, RegistryValueKind.DWord);

							startupFlags = (StartupFlags)uiFlags;

							//Mark it that we've got something
							bRes = true;
						}
					}
				}
			}
			catch (Exception ex)
			{
				LogDiagnosticSpecError(1092, ex);
				bRes = false;
			}

			return bRes;
		}


		private void UpgradeConfigIfNeeded()
		{
			//Checks if configuration file needs to be upgraded, and if so, does the upgrade
			//INFO: The placement of configuration file is tricky, as it can change with new versions of Outlook.
			string strConfigPath = "";

			try
			{
				//Get config path
				strConfigPath = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal).FilePath;

				//See if we need to import data from a file after initial installation
				string strImportFilePath;
				if (!GetInitialImportData(out strImportFilePath, out gStartupFlgs))
				{
					//Notmal start - without initial parameters from MSI


					//See if our setting for an upgrage is on
					//INFO: It will be on if settings were loaded from scratch (or anew)
					if (Properties.Settings.Default.OutlookHdrsUpgrade)
					{
						//Make sure that the new settings file exists
						Properties.Settings.Default.OutlookHdrsUpgrade = false;
						Properties.Settings.Default.Save();



						string strOldConfigFilePath = "";

						//Example of what we're doing here:
						//
						//Going from:
						//		C:\Users\Admin\AppData\Local\Microsoft_Corporation\OutlookHeaders.vsto_vstol_Path_i5vznyicnelzuytw513coiw3qpofvydj\12.0.4518.1014\user.config
						//To:
						//		C:\Users\Admin\AppData\Local\Microsoft_Corporation\OutlookHeaders.vsto_vstol_Path_3k1kd4rpcshhqmybeo235ri5audy31kl\16.0.4266.1003\user.config
						//
						DirectoryInfo configDir = new FileInfo(strConfigPath).Directory;
						if (configDir != null)
						{
							//Collect places to search for our config file
							List<string> arrCheckFldrs = new List<string>();

							//Go up one level
							//INFO: This is for minor updates of Outlook
							DirectoryInfo parDir = configDir.Parent;
							string strExcludeDir = configDir.Name;

							if (parDir.Exists)
							{
								foreach (DirectoryInfo dir in parDir.GetDirectories())
								{
									//Exclude our dir
									if (string.Compare(strExcludeDir, dir.Name, true) != 0)
									{
										//Add it to search for our config file
										arrCheckFldrs.Add(dir.FullName);
									}
								}
							}

							//Go up two levels
							//INFO: This is for major updates of Outlook
							strExcludeDir = parDir.Name;
							parDir = parDir.Parent;

							if (parDir.Exists)
							{
								//Get this project name
								string strProjName = getThisProjectName();
								if (!string.IsNullOrEmpty(strProjName))
								{
									foreach (DirectoryInfo dir in parDir.GetDirectories())
									{
										//Exclude our dir
										if (string.Compare(strExcludeDir, dir.Name, true) != 0)
										{
											//Look for anything that starts with the name of our solution
											if (dir.Name.StartsWith(strProjName + ".", StringComparison.InvariantCultureIgnoreCase))
											{
												//Check everything in this folder (one level deep)
												foreach (DirectoryInfo dir2 in dir.GetDirectories())
												{
													//Add it to search for our config file
													arrCheckFldrs.Add(dir2.FullName);
												}
											}
										}
									}
								}
								else
									ThisAddIn.LogDiagnosticSpecWarning(1055);
							}

							//Get our config file name
							string strUserConfigFileName = Path.GetFileName(strConfigPath);

							//Go through all of the config files and pick the latest one by its last-write date
							//INFO: It is our (not very bullet-proof) assumption that such will be the latest config file.
							DateTime dtmLast = DateTime.MinValue;

							foreach (string strFldrPath in arrCheckFldrs)
							{
								string strUserConfigPath = strFldrPath + Path.DirectorySeparatorChar + strUserConfigFileName;

								FileInfo fiUserConfig = new FileInfo(strUserConfigPath);
								if (fiUserConfig != null && fiUserConfig.Exists)
								{
									DateTime dtm = fiUserConfig.LastWriteTimeUtc;
									if (string.IsNullOrEmpty(strOldConfigFilePath) ||
										dtm > dtmLast)
									{
										dtmLast = dtm;
										strOldConfigFilePath = strUserConfigPath;
									}
								}
							}

							//Did we find a previous config file?
							if (!string.IsNullOrEmpty(strOldConfigFilePath))
							{
								//Then try to copy it in place of our current file (allow overwrites)
								File.Copy(strOldConfigFilePath, strConfigPath, true);



							}
							else
							{
								//Issue a warning that we haven't found such a file
								//NOTE that this will also happen during the first installation of our add-in.
								ThisAddIn.LogDiagnosticSpecWarning(1053);
							}
						}


						//Call upgrade method
						//INFO: TBH, I'm not sure what it does. But documentation suggests calling it ...
						Properties.Settings.Default.Upgrade();

						//Reset the setting so that this method doesn't get called again
						Properties.Settings.Default.OutlookHdrsUpgrade = false;

						//And save everything in persistent storage
						Properties.Settings.Default.Save();


						//Mark in the event log that we just upgraded
						ThisAddIn.LogDiagnosticEvent(1051, "Upgrading settings.\n" +
							"New config path: " + strConfigPath + "\n" +
							"Old config path: " + strOldConfigFilePath + "\n" +
							"Installer type: " + ThisAddIn.getCurrentInstallerType().ToString() + "\n",
							EventLogEntryType.Information);
					}
				}
				else
				{
					//First installation -- using imported parameters from an MSI installer

					//Set initial enabled state & logging-when-sending-emails flag
					Properties.Settings.Default.OutlookHdrsEnabled = (gStartupFlgs & StartupFlags.STUP_FLG_START_DISABLED) != 0 ? false : true;
					Properties.Settings.Default.OutlookHdrsLogWhenSending = (gStartupFlgs & StartupFlags.STUP_FLG_LOG_WHEN_SENDING_EMAILS) != 0;


					//Do we have an import data file
					if (!string.IsNullOrEmpty(strImportFilePath))
					{
						//If so, then it's been already validated
						try
						{
							//Read data from the file
							string strXml = System.IO.File.ReadAllText(strImportFilePath);

							Properties.Settings.Default.OutlookHdrsMailHdrs = strXml;
						}
						catch (Exception ex)
						{
							//Failed
							LogDiagnosticSpecError(1097, "Failed to import file: " + strImportFilePath + " > " + ex.ToString());
						}

						//Put a note
						LogDiagnosticEvent(1096, "Imported the following data file: \"" + strImportFilePath + "\"", EventLogEntryType.Information);
					}


					//Reset upgrade setting, as we don't need it after an import
					Properties.Settings.Default.OutlookHdrsUpgrade = false;

					//And save everything in persistent storage
					Properties.Settings.Default.Save();
				}



				//Update Registry settings
				UpdateThisAppRegistrySettings(strConfigPath);
			}
			catch (Exception ex)
			{
				//Failed
				ThisAddIn.LogDiagnosticSpecError(1050, "config: " + strConfigPath + " > " + ex.ToString());
				SystemSounds.Exclamation.Play();
			}
		}

		private void UpgradeConfigAppVersionIfNeeded()
		{
			//Checks if this app version requires an upgrade from a previous version
			//INFO: Does nothing if version is the same
			//IMPORTANT: Must be called after UpgradeConfigIfNeeded() call!
			try
			{
				// --- for now it's an empty method. No need to do anything here yet ---


				//Lastly write current version into the settings & save
				Properties.Settings.Default.OutlookHdrsVer = ThisAddIn.gAppVersion.getVersion();
				Properties.Settings.Default.Save();
			}
			catch (Exception ex)
			{
				//Failed
				ThisAddIn.LogDiagnosticSpecError(1052, ex);
				SystemSounds.Exclamation.Play();
			}
		}


		private void UpdateThisAppRegistrySettings(string strConfigPath)
		{
			//Write into registry this app's settings
			//'strConfigPath' = path to user.config file
			try
			{
				using (RegistryKey regK = Registry.CurrentUser.CreateSubKey(GlobalDefs.GlobDefs.gkstrAppRegistryKey))
				{
					//Set static values
					regK.SetValue(GlobalDefs.GlobDefs.gkstrAppRegVal_Version, ThisAddIn.gAppVersion.getVersion(), RegistryValueKind.String);
					regK.SetValue(GlobalDefs.GlobDefs.gkstrAppRegVal_InstallFldr, AppDomain.CurrentDomain.BaseDirectory, RegistryValueKind.String);

					//Open all existing config paths
					string[] arrCFPs = (string[])regK.GetValue(GlobalDefs.GlobDefs.gkstrAppRegVal_ConfigFilePaths, new string[] { }, RegistryValueOptions.None);

					//See if we have ours already
					bool bFoundIt = false;
					foreach (string strPath in arrCFPs)
					{
						if (string.Compare(strPath, strConfigPath, true) == 0)
						{
							//Got it
							bFoundIt = true;
							break;
						}
					}

					if (!bFoundIt)
					{
						//Add our config path and save it
						List<string> arr = new List<string>(arrCFPs);
						arr.Add(strConfigPath);

						//Save it in registry
						regK.SetValue(GlobalDefs.GlobDefs.gkstrAppRegVal_ConfigFilePaths, arr.ToArray(), RegistryValueKind.MultiString);
					}
				}
			}
			catch (Exception ex)
			{
				//Failed
				LogDiagnosticSpecError(1070, ex);
			}
		}


		public enum EmlMsgType
		{
			EMT_Reply,
			EMT_ReplyAll,
			EMT_Forward,
		}

		public class EmlMsgFmtItem
		{
			public EmlMsgType type;
			public Outlook.MailItem objMailItm;
			public Microsoft.Office.Interop.Outlook.Inspector objInsp;
		}

		private static void adjustMessageFormatting(object Item, EmlMsgType msgType)
		{
			//Adjust formatting in a new message (if needed)
			//'Item' = email item being replied or forwarded
			//'msgType' = type of operation
			try
			{
				Outlook.MailItem objMailItem = (Outlook.MailItem)Item;
				Microsoft.Office.Interop.Outlook.Inspector insp = objMailItem.GetInspector;
				if (insp != null)
				{
					EmlMsgFmtItem emfi = new EmlMsgFmtItem() { objMailItm = objMailItem, objInsp = insp, type = msgType };

					//Need to do the change after a short delay
					//INFO: Don't know why.... my guess is that at this moment the command bar is not fully loaded...
					//      This is an awful way to program it, but coupled with the fact that all this stuff
					//      is horribly documented, I have no other choice but to leave it like this :(
					System.Windows.Forms.Timer tmr = new Timer();
					tmr.Interval = 500;
					tmr.Tick += new EventHandler(onTimer_CheckAndChangeMsgFormat);
					tmr.Tag = emfi;
					tmr.Start();

				}
				else
				{
					//No inspector
					LogDiagnosticSpecError(1137, "type=" + msgType.ToString());
				}
			}
			catch (Exception ex)
			{
				//Exception
				LogDiagnosticSpecError(1135, "type=" + msgType.ToString() + " > " + ex.ToString());
			}

		}

		private static void onTimer_CheckAndChangeMsgFormat(object sender, EventArgs e)
		{
			//Called after a short delay
			EmlMsgFmtItem emfi = null;
			string strEmailFrom = "";
			string strErrDesc;

			try
			{
				//Stop timer from running - we need it only to fire once!
				System.Windows.Forms.Timer tmr = (System.Windows.Forms.Timer)sender;
				tmr.Stop();

				//Parameters from the caller
				emfi = (EmlMsgFmtItem)tmr.Tag;



				//Read config
				bool bLog = Properties.Settings.Default.OutlookHdrsLogWhenSending;
				bool bEnabled = Properties.Settings.Default.OutlookHdrsEnabled;

				//Only if add-in is enabled
				if (bEnabled)
				{
					//Get email details
					strEmailFrom = emfi.objMailItm.SendUsingAccount.SmtpAddress;

					FormDefineHeaders.OutboundEmailInfo emlFmt = new FormDefineHeaders.OutboundEmailInfo();


					//Get data to process
					XmlDocument xmlHdrs = readHeaders(out strErrDesc);
					if (xmlHdrs != null)
					{
						XmlElement xmlRoot = xmlHdrs.DocumentElement;
						if (xmlRoot == null)
							new Exception("[1139]");

						//Find the <ALL> node first
						XmlNode xNodeAll = xmlRoot.SelectSingleNode(FormDefineHeaders.kgStrNodeNm_Account + "[@" + FormDefineHeaders.kgStrNodeAttNm_All + "]");
						if (xNodeAll == null)
							new Exception("[1140]");	//We must have ALL node!

						//Process for the <ALL> node
						processNodeHeadersForFormatChange(xNodeAll, ref emlFmt);


						XmlAttribute xAtt;

						//Then go through the other nodes and match them to our account
						foreach (XmlNode xNode in xmlRoot.ChildNodes)
						{
							//Skip <ALL> node
							if (xNode != xNodeAll)
							{
								xAtt = xNode.Attributes[FormDefineHeaders.kgStrNodeAttNm_Name];
								string strName = (xAtt != null) ? xAtt.Value : null;

								if (!string.IsNullOrEmpty(strName) &&
									string.Compare(strName, strEmailFrom, true) == 0)
								{
									//Match -- process these headers
									processNodeHeadersForFormatChange(xNode, ref emlFmt);

									break;
								}
							}
						}


						//See if we need to convert?
						string strTxt;
						bool bConverted = false;
						if (emlFmt.sendEmlAs == FormDefineHeaders.EmailFmt.EML_FMT_SendAsHTML)
						{
							//Convert to HTML
							emfi.objInsp.CommandBars.ExecuteMso("MessageFormatHtml");

							if (emlFmt.useRemoveTrackingPixel == FormDefineHeaders.TriState.TS_True)
							{
								//Remove tracking pixel
								strTxt = removeTrackingPixel(emfi.objMailItm.Body, emfi.objMailItm.BodyFormat, emlFmt.strTrackingDomains);
							}
							else
							{
								//Use text as-is
								strTxt = emfi.objMailItm.Body;
							}

							//And convert to HTML from plaintext
							//INFO: We need to do this to compensate for the Outlook bug that includes double-spaces between paragraphs
							emfi.objMailItm.HTMLBody = convertPlaintextToHtml(strTxt, emlFmt);
							//emfi.objMailItm.Body = strTxt;

							bConverted = true;
						}
						else if (emlFmt.sendEmlAs == FormDefineHeaders.EmailFmt.EML_FMT_SendAsText)
						{
							//Convert to plaintext
							emfi.objInsp.CommandBars.ExecuteMso("MessageFormatPlainText");	//"MessageFormatRichText");

							if (emlFmt.useRemoveTrackingPixel == FormDefineHeaders.TriState.TS_True)
							{
								//Remove tracking pixel
								emfi.objMailItm.Body = removeTrackingPixel(emfi.objMailItm.Body, emfi.objMailItm.BodyFormat, emlFmt.strTrackingDomains);
							}

							//emfi.objMailItm.HTMLBody = "";

							bConverted = true;
						}


						if (bConverted)
						{
							if (bLog)
							{
								LogDiagnosticEvent(1141, "Converted msg from " + emfi.objMailItm.BodyFormat.ToString() + 
									" to " + emlFmt.sendEmlAs.ToString() + " for " + emfi.type.ToString() + "\n" + 
									"OverwriteMailClient=" + FormDefineHeaders.convertTriStateToDebugStr(emlFmt.useOverwriteMailClient) + " > \"" + emlFmt.strMailClientName + "\"\n" +
									"SuppressReceivedHdr=" + FormDefineHeaders.convertTriStateToDebugStr(emlFmt.useSuppressRcvdHeader) + " > Tmpl=\"" + emlFmt.strRcvdHeader + "\"\n" +
									"SetCSS=" + FormDefineHeaders.convertTriStateToDebugStr(emlFmt.useCssStyles) + " > \"" + emlFmt.strCssStyles + "\"\n" +
									"TrackingPixel=" + FormDefineHeaders.convertTriStateToDebugStr(emlFmt.useRemoveTrackingPixel) + " > \"" + emlFmt.strTrackingDomains + "\"\n"
									, EventLogEntryType.Information);
							}
						}
					}
					else
					{
						//Failed to read headers - rethrow
						new Exception("[1138] " + strErrDesc);
					}
				}


			}
			catch (Exception ex)
			{
				//Exception
				LogDiagnosticSpecError(1136, "type=" + (emfi != null ? emfi.type.ToString() : "<null>") + ", emailFrom=" + strEmailFrom + " > " + ex.ToString());
			}
		}

		private static void processNodeHeadersForFormatChange(XmlNode xNode, ref FormDefineHeaders.OutboundEmailInfo emlFmt)
		{
			//Process 'xNode' node for format change
			//'xNode' = XML node with config data with headers for this account
			//'emlFmt' = will be updated if needed by the node

			FormDefineHeaders.OutboundEmailInfo efi = FormDefineHeaders.getOutboundEmlInfoFromNode(xNode);

			if (efi.useOverwriteMailClient != FormDefineHeaders.TriState.TS_Inherit)
			{
				emlFmt.useOverwriteMailClient = efi.useOverwriteMailClient;
				emlFmt.strMailClientName = efi.strMailClientName;
			}

			if (efi.useSuppressRcvdHeader != FormDefineHeaders.TriState.TS_Inherit)
			{
				emlFmt.useSuppressRcvdHeader = efi.useSuppressRcvdHeader;
				emlFmt.strRcvdHeader = efi.strRcvdHeader;
			}

			if (efi.sendEmlAs != FormDefineHeaders.EmailFmt.EML_FMT_None)
			{
				emlFmt.sendEmlAs = efi.sendEmlAs;
			}

			if (efi.useCssStyles != FormDefineHeaders.TriState.TS_Inherit)
			{
				emlFmt.useCssStyles = efi.useCssStyles;
				emlFmt.strCssStyles = efi.strCssStyles;
			}

			if (efi.useRemoveTrackingPixel != FormDefineHeaders.TriState.TS_Inherit)
			{
				emlFmt.useRemoveTrackingPixel = efi.useRemoveTrackingPixel;
				emlFmt.strTrackingDomains = efi.strTrackingDomains;
			}
		}

		public static HashSet<string> getDomainListFromStr(string strDomains)
		{
			//'strDomains' = list of space separated domain names
			//RETURN:
			//		= List of domains from 'strDomains'
			HashSet<string> resSet = new HashSet<string>();

			if (string.IsNullOrEmpty(strDomains))
				strDomains = "";

			//Split into parts
			string[] arrDoms = strDomains.Split(new char[] {' ', '\t', '\r', '\n'});
			string strD;

			foreach(string strDom in arrDoms)
			{
				strD = strDom.Trim();
				if(!string.IsNullOrEmpty(strD))
				{
					resSet.Add(strD);
				}
			}

			return resSet;
		}

		public static string convertPlaintextToHtml(string strTxt, FormDefineHeaders.OutboundEmailInfo emlFmt)
		{
			//Convert plaintext message from 'strTxt' to HTML format
			//'emlFmt' = formatting details
			//RETURN:
			//		= HTML formatted message
			StringBuilder sbRes = new StringBuilder();

			sbRes.Append("<html><head>");

			if(emlFmt.useCssStyles == FormDefineHeaders.TriState.TS_True)
				sbRes.Append("<style type=\"text/css\">" + emlFmt.strCssStyles + "</style>");

			sbRes.Append("</head><body><p>");

			int nLn = strTxt.Length;
			for (int i = 0; i < nLn; i++)
			{
				char c = strTxt[i];

				if (c == '&')
				{
					sbRes.Append("&amp;");
				}
				else if (c == '<')
				{
					sbRes.Append("&lt;");
				}
				else if (c == '>')
				{
					sbRes.Append("&gt;");
				}
				else if (c == '"')
				{
					sbRes.Append("&quot;");
				}
				else if (c == '\r')
				{
					//See if \n follows it
					if (i + 1 < nLn && strTxt[i + 1] == '\n')
					{
						//Skip it
						continue;
					}
					else
					{
						sbRes.Append("<br>");
					}
				}
				else if (c == '\n')
				{
					sbRes.Append("<br>");
				}
				else if (char.IsWhiteSpace(c))
				{
					//Make sure that new-line doesn't follow it
					bool bSkip = false;
					int j = i + 1;
					for (; j < nLn; j++)
					{
						char z = strTxt[j];
						if (z == '\r' || z == '\n')
						{
							//Skip all these whitespaces
							i = j - 1;
							bSkip = true;

							break;
						}
						else if (!char.IsWhiteSpace(z))
						{
							break;
						}
					}

					if (!bSkip)
					{
						if (c != '\t')
							sbRes.Append(c);
						else
							sbRes.Append("   ");
					}
				}
				else
					sbRes.Append(c);
			}

			sbRes.Append("</p></body></html>");

			return sbRes.ToString();
		}


		public static string removeTrackingPixel(string strTxt, Outlook.OlBodyFormat msgFmt, string strDomains)
		{
			//Remove tracking pixel from 'strTxt'
			//'msgFmt' = format of text in 'strTxt'
			//'strDomains' = space-separated list of domains to remove tracking pixel for
			//RETURN:
			//		= Updated text
			StringBuilder sbRes = new StringBuilder();

			//List of domains providing a tracking pixel
			HashSet<string> arrDoms = getDomainListFromStr(strDomains);

			string strTrackUrl, strHost;
			int nFndTracking;

			int nLn = strTxt.Length;
			for (int i = 0; i < nLn; i++)
			{
				char c = strTxt[i];

				// <http://ntt5k4sd.r.us-west-2.awstrack.me/I0/blah-blah>
				if (msgFmt == Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatHTML)
				{
					// &lt;   or   <  
					if(c == '<')
						nFndTracking = i + 1;
					else
						nFndTracking = i + 3 < nLn && c == '&' && strTxt[i + 1] == 'l' && strTxt[i + 2] == 't' && strTxt[i + 3] == ';' ? i + 4 : 0;
				}
				else
				{
					// <
					nFndTracking = c == '<' ? i + 1 : 0;
				}

				if(nFndTracking > 0)
				{
					int nEndTracking = -1;
					int nSet_i = 0;

					for (int j = nFndTracking; j < nLn; j++)
					{
						char z = strTxt[j];

						if (z == '>')
						{
							nEndTracking = j;
							nSet_i = j;
							break;
						}

						if (msgFmt == Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatHTML)
						{
							//  &gt;
							if (j + 3 < nLn && z == '&' && strTxt[j + 1] == 'g' && strTxt[j + 2] == 't' && strTxt[j + 3] == ';')
							{
								nEndTracking = j;
								nSet_i = j + 3;
								break;
							}
						}
					}

					if (nEndTracking > nFndTracking)
					{
						bool bIsTrackingURL = false;
						strTrackUrl = strTxt.Substring(nFndTracking, nEndTracking - nFndTracking);

						try
						{
							//Does it begin with http: or https:
							if (strTrackUrl.StartsWith("http:", StringComparison.CurrentCultureIgnoreCase) ||
								strTrackUrl.StartsWith("https:", StringComparison.CurrentCultureIgnoreCase))
							{
								Uri url = new Uri(strTrackUrl);
								strHost = url.Host;

								foreach (string strDom in arrDoms)
								{
									if (strHost.EndsWith(strDom, StringComparison.CurrentCultureIgnoreCase))
									{
										//This is a tracking pixel
										bIsTrackingURL = true;

										break;
									}
								}
							}
						}
						catch (Exception ex)
						{
							//Failed with this domain name
							LogDiagnosticSpecError(1134, "url=\"" + strTrackUrl + "\" > " + ex.ToString());
						}

						if (bIsTrackingURL)
						{
							//Found tracking URL - then skip it
							i = nSet_i;
							continue;
						}
					}
				}

				sbRes.Append(c);
			}

			return sbRes.ToString();
		}

		public static bool openWebPage(string strURL)
		{
			//Open 'strURL' in a web page
			//RETURN:
			//		= true if success
			return openWebPage(strURL, true);
		}

		public static bool openWebPage(string strURL, bool bAllowErrorMsgBox)
		{
			//Open 'strURL' in a web page
			//'bAllowErrorMsgBox' = true to allow to show error message box if failed to open
			//RETURN:
			//		= true if success
			bool bRes = false;

			try
			{
				System.Diagnostics.Process.Start(strURL);

				bRes = true;
			}
			catch (Exception ex)
			{
				//Exception
				LogDiagnosticSpecError(1143, "URL: " + strURL + " > " + ex.ToString());
				bRes = false;

				if(bAllowErrorMsgBox)
				{
					//Show user error
					MessageBox.Show("ERROR: Failed to show the following URL.\n\n" + strURL, GlobalDefs.GlobDefs.gkstrAppNameFull, MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}

			return bRes;
		}


		public enum InstallerType
		{
			INST_Unknown = 0,
			INST_PerUser = 1,
			INST_AllUsers = 2,
			INST_Bundled = 3,		//Also "for all-users" but installed in a Bootstrapper bundle
		}

		public static InstallerType getCurrentInstallerType()
		{
			//RETURN:
			//		= The type of installer that was used to install this add-in

			//These values come from the MSI
			string strRegKey = @"SOFTWARE\Microsoft\Office\Outlook\Addins\" + GlobalDefs.GlobDefs.gkstrAppName;
			string strRegValUsrType = "InstallerUsrType";		//REG_DWORD: 1=per-user, 2=all users
			string strRegValBundle = "InstallerBundle";			//REG_SZ: 1 if bundle installation

			try
			{
				//Try local user first
				using (RegistryKey regCU = Registry.CurrentUser.OpenSubKey(strRegKey, false))
				{
					if (regCU != null)
					{
						object objV = regCU.GetValue(strRegValUsrType);
						if (objV != null)
						{
							//Its type is REG_DWORD
							int nV = (int)objV;
							if (nV == 1)
							{
								//It's per-user install
								return InstallerType.INST_PerUser;
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				//Error
				ThisAddIn.LogDiagnosticSpecError(1156, ex);
			}

			try
			{
				//Try local machine next
				using (RegistryKey regLM = Registry.LocalMachine.OpenSubKey(strRegKey, false))
				{
					if (regLM != null)
					{
						object objV = regLM.GetValue(strRegValUsrType);
						if (objV != null)
						{
							//Its type is REG_DWORD
							int nV = (int)objV;
							if(nV == 2)
							{
								//Get bundle value
								objV = regLM.GetValue(strRegValBundle);
								if (objV != null)
								{
									//It's type is REG_SZ
									if (int.TryParse((string)objV, out nV))
									{
										if (nV == 1)
										{
											//This ia bundled install
											return InstallerType.INST_Bundled;
										}
									}
								}

								//Otherwise it's just for all users
								return InstallerType.INST_AllUsers;
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				//Error
				ThisAddIn.LogDiagnosticSpecError(1157, ex);
			}

			return InstallerType.INST_Unknown;
		}


		public static bool LogDiagnosticEvent(int nSpecErr, string strMsg, EventLogEntryType type)
		{
			//Places a diagnostic message into an event log
			//INFO: This function will add current time & version of the app
			//'nSpecErr' = special unique error ID in the source code. To generate it, use SeqIDGen: https://dennisbabkin.com/seqidgen
			//'strMsg' = additional text message to log to describe an error
			//'type' = type of a message: error, warning, information
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
						": [" + GlobalDefs.GlobDefs.gkstrAppNameFull + " v." + ThisAddIn.gAppVersion.getVersion() + "]" +
						"[" + (isThis64bitProc() ? "x64" : "x86") + ">" + Environment.UserName + "]" +
						"[" + nSpecErr + "]" +
						(!string.IsNullOrEmpty(strMsg) ? " " : "") + strMsg,
						type, 1, 0);

					bRes = true;
				}
			}
			catch (Exception)
			{
				//Failed - no logging here
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
			//'ex' = exception with the error description to log
			//RETURN:
			//		= true if logged without any errors
			return LogDiagnosticEvent(nSpecErr, ex.ToString(), EventLogEntryType.Error);
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
