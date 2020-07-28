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
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Xml;
using System.IO;
using System.Media;
using System.Diagnostics;
using System.Collections.Specialized;



namespace OutlookHeaders
{
    public partial class FormDefineHeaders : Form
    {
		public enum AccountItmType
		{
			ACCT_ITP_DEFAULT,
			ACCT_ITP_ALL,
		}

        public class AccountItem
        {
            public string strAcctName { get; set; }
			public AccountItmType type { get; set; }

            public override string ToString()
            {
                return type != AccountItmType.ACCT_ITP_ALL ? strAcctName : "<ALL>";
            }

			public string toDebugString()
			{
				string str = "Type=" + (type == AccountItmType.ACCT_ITP_ALL ? "ALL" : "Regular");
				if (type != AccountItmType.ACCT_ITP_ALL)
				{
					str += ", Acct=" + strAcctName;
				}

				return str;
			}
        }

		public class HeaderItem
		{
			public ListViewItem lvItem = null;			//List view item
			public XmlNode xmlNode = null;				//XML node with cached data
			public bool bAllItem = false;				//true if this is ALL item
		}

		public class ItemInfo
		{
			public string strName = "";
			public string strVal = "";
			public bool bEnabled = true;
		}

		public enum EmailFmt
		{
			EML_FMT_None = 0,				//No value was used
			EML_FMT_Dont_Change,			//Don't change how emails are sent out
			EML_FMT_SendAsText,				//Send all emails as plain text
			EML_FMT_SendAsHTML,				//Send all emails as HTML

			EML_FMT_Count					//USE LAST!
		}

		public enum TriState
		{
			TS_False = 0,
			TS_True = 1,
			TS_Inherit = 2,
		}

		public class OutboundEmailInfo
		{
			public EmailFmt sendEmlAs = EmailFmt.EML_FMT_Dont_Change;		//How to force sending emails
			public TriState useOverwriteMailClient = TriState.TS_Inherit;	//true to overwrite mail client name
			public string strMailClientName = "";							//[Depends on 'useOverwriteMailClient'] mail client name to overwrite with
			public TriState useSuppressRcvdHeader = TriState.TS_Inherit;	//true to suppress "Received" header
			public string strRcvdHeader = "";								//[Depends on 'useSuppressRcvdHeader'] "Received" header to use
			public TriState useCssStyles = TriState.TS_Inherit;				//true to use 'strCssStyles'
			public string strCssStyles = "";								//Contents of the <style></style> tag for the email for HTML format
			public TriState useRemoveTrackingPixel = TriState.TS_Inherit;	//true to remove tracking pixels for domains in 'strTrackingDomains'
			public string strTrackingDomains = "";							//[Depends on 'useRemoveTrackingPixel'] space-separated domain names
		}

		public class EmlFmt_Item
		{
			public EmailFmt emlFmt;
			public string strName;

			public override string ToString()
			{
				return strName;
			}
		}

		public class ImportedDoc
		{
			public XmlDocument xmlDoc = null;				//XML data read
			public string strVersion = "";					//Version of the file where 'xmlDoc' was read from
		}


        public ThisAddIn _this = null;
		XmlDocument _xmlHdrs = null;						//Current XML document for the account headers
		XmlNode _xmlNodeSelAcct = null;						//XML node for currently selected account, or null otherwise
		AccountItem _acctItem = null;						//Currently selected account item, or null otherwise
		bool _bAllowCatchCmbAcctsChanges = false;			//true to allow to catch changes in the 'comboBoxAccounts' combo box
		bool _bAllowCatchListHdrsChanges = false;			//true to allow to catch changes in the 'listHdrs' list
		bool _gbDirty = false;								//true if data was modified since it was loaded


		//XML node names and attributes in the saved config file:
		public static string kgStrNodeNm_Root = "Accounts";
		public static string kgStrNodeAttNm_RootApp = "App";
		public static string kgStrNodeAttNm_RootVer = "Ver";

		public static string kgStrNodeNm_Account = "account";
		public static string kgStrNodeAttNm_All = "all";
		public static string kgStrNodeAttNm_Name = "name";
		public static string kgStrNodeAttNm_Enabled = "enabled";
		public static string kgStrNodeAttNm_OverwriteMailClientNm = "ovwrtmailclnt";	//0=no, 1=yes, 2=inherit
		public static string kgStrNodeAttNm_MailClientNm = "mailclnt";
		public static string kgStrNodeAttNm_SuppressRcvdHdr = "sprsrcvdhdr";	//0=no, 1=yes, 2=inherit
		public static string kgStrNodeAttNm_RcvdHdr = "rcvdhdr";
		public static string kgStrNodeAttNm_SendAs = "sendas";
		public static string kgStrNodeAttNm_SetCSSStyles = "setcss";			//0=no, 1=yes, 2=inherit
		public static string kgStrNodeAttNm_CSSStyles = "css";
		public static string kgStrNodeAttNm_RemTrackPixel = "remtrkpxl";		//0=no, 1=yes, 2=inherit
		public static string kgStrNodeAttNm_TrackDoms = "trkdom";

		public static string kgStrNodeNm_Header = "header";


        public FormDefineHeaders()
        {
            InitializeComponent();
        }

        private void FormDefineHeaders_Load(object sender, EventArgs e)
        {
            //When form is loading
			string strErrDesc;

			//Restore window location
			Size szWnd = Properties.Settings.Default.OutlookHdrsWndSize;

			//Only if saved coordinates are valid
			if (szWnd.Width > 0 && szWnd.Width >= this.MinimumSize.Width &&
				szWnd.Height > 0 && szWnd.Height >= this.MinimumSize.Height)
			{
				//Read position
				Point pntWnd = Properties.Settings.Default.OutlookHdrsWndLocation;

				//Get working rect of the monitor we're in
				Rectangle rcWrk = Screen.FromControl(this).WorkingArea;

				//Calculate where title area for our window is
				int nTitleBarH = this.RectangleToScreen(this.ClientRectangle).Top - this.Top;
				Rectangle rcWndTitle = new Rectangle(pntWnd.X, pntWnd.Y, szWnd.Width, nTitleBarH);

				//Use screen's work area that has been somewhat shrunk (so that the user can see our window, in case it's shown on the very edge)
				Rectangle rcUseWrk = rcWrk;
				rcUseWrk.Inflate(-(int)(rcWrk.Width * 0.025), -(int)(rcWrk.Height * 0.037));

				//Is this point somewhere within the monitor of our original window?
				if (rcUseWrk.IntersectsWith(rcWndTitle))
				{
					//Make sure the size of the window is not too large
					Size szMaxAllowed = new Size(rcWrk.Width + 40, rcWrk.Height + 40);		//Why 40? Well, why not!

					if (szWnd.Width > szMaxAllowed.Width)
						szWnd.Width = szMaxAllowed.Width;

					if (szWnd.Height > szMaxAllowed.Height)
						szWnd.Height = szMaxAllowed.Height;


					//Use these coordinates
					Location = pntWnd;
					Size = szWnd;
					WindowState = Properties.Settings.Default.OutlookHdrsWndMax ? FormWindowState.Maximized : FormWindowState.Normal;
				}
			}


			//Limit text size in edit controls
			textBoxMailClientNm.MaxLength = 1024;
			textBoxSuppressRcvd.MaxLength = 1024;
			textBoxCssStyle.MaxLength = 4096;
			textBoxTrckDomains.MaxLength = 1024;


			//Load previous headers from settings
			_xmlHdrs = ThisAddIn.readHeaders(out strErrDesc);
			if (_xmlHdrs != null)
			{
				//Reload UI
				reloadAllAcoounts();
			}
			else
			{
				//Error
				ThisAddIn.LogDiagnosticSpecError(1007, strErrDesc);
				MessageBox.Show("ERROR: Failed to read saved XML data.\n\n" + strErrDesc, GlobalDefs.GlobDefs.gkstrAppNameFull, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}


			//Get default sizes of list columns
			int nW0, nW1, nW2;
			getDefaultHeaderListColumnSizes(out nW0, out nW1, out nW2);

			//See if we have previously saved column sizes
			if (Properties.Settings.Default.OutlookHdrsWndLVColSzs != null &&
				Properties.Settings.Default.OutlookHdrsWndLVColSzs.Count == 3)
			{
				//Read sizes
				int w0 = 0, w1 = 0, w2 = 0;
				if (int.TryParse(Properties.Settings.Default.OutlookHdrsWndLVColSzs[0].ToString(), out w0) &&
					int.TryParse(Properties.Settings.Default.OutlookHdrsWndLVColSzs[1].ToString(), out w1) &&
					int.TryParse(Properties.Settings.Default.OutlookHdrsWndLVColSzs[2].ToString(), out w2))
				{
					//Get max allowed width
					int nMaxW = listHdrs.ClientSize.Width + 50;		//Why 50? Why not.
					if (nMaxW < 50)
						nMaxW = 50;

					if (w0 > 0 && w0 < nMaxW &&
						w1 > 0 && w1 < nMaxW &&
						w2 > 0 && w2 < nMaxW)
					{
						//Use these sizes
						nW0 = w0;
						nW1 = w1;
						nW2 = w2;
					}
				}
			}

			//Add list columns (no localization here for speed -- just hard-coding string literals)
			listHdrs.Columns.Add("<ALL>", nW0, HorizontalAlignment.Left);
			listHdrs.Columns.Add("Field Name", nW1, HorizontalAlignment.Left);
			listHdrs.Columns.Add("Field Body", nW2, HorizontalAlignment.Left);


			//Set "not-dirty" state originally
			setDirtyState(false);
        }


		private void reloadAllAcoounts()
		{
			//Reload UI for all accounts
			//INFO: It usually has to be done after a radical change in the window ...

			//Reset global flags
			_bAllowCatchCmbAcctsChanges = false;
			_xmlNodeSelAcct = null;

			try
			{
				Microsoft.Office.Interop.Outlook.Accounts acts = _this.Application.Session.Accounts;

				int nCntActs = acts.Count;

				//Remove all existing items
				comboBoxAccounts.Items.Clear();

				//Remove all items in the list too
				listHdrs.Items.Clear();


				//Always add "All" item
				comboBoxAccounts.Items.Add(new AccountItem() { type = AccountItmType.ACCT_ITP_ALL });

				//Microsoft are special -- they are counting indeces from 1 :)
				for (int i = 1; i <= nCntActs; i++)
				{
					comboBoxAccounts.Items.Add(new AccountItem() { strAcctName = acts[i].DisplayName, type = AccountItmType.ACCT_ITP_DEFAULT });
				}


				//Set initial XML data
				XmlElement xRoot = _xmlHdrs.DocumentElement;
				if (xRoot == null)
				{
					//No root, add it
					XmlDeclaration xmlDec = _xmlHdrs.CreateXmlDeclaration("1.0", "utf-8", null);
					_xmlHdrs.InsertBefore(xmlDec, _xmlHdrs.DocumentElement);

					//Add root
					xRoot = _xmlHdrs.CreateElement(kgStrNodeNm_Root);

					//Add version of the file
					XmlAttribute xAttApp = _xmlHdrs.CreateAttribute(kgStrNodeAttNm_RootApp);
					xAttApp.Value = GlobalDefs.GlobDefs.gkstrAppNameFull;
					xRoot.Attributes.Append(xAttApp);

					XmlAttribute xAttVer = _xmlHdrs.CreateAttribute(kgStrNodeAttNm_RootVer);
					xAttVer.Value = ThisAddIn.kstrDataFileVer;		//Special data file version
					xRoot.Attributes.Append(xAttVer);

		
					_xmlHdrs.AppendChild(xRoot);
				}

				string strAcctName;
				XmlNode xNode;
				int nCnt = comboBoxAccounts.Items.Count;

				//Add nodes for accounts that we didn't have
				for (int i = 0; i < nCnt; i++)
				{
					//See if we have our account nodes
					AccountItem actItm = (AccountItem)comboBoxAccounts.Items[i];

					if (actItm.type != AccountItmType.ACCT_ITP_ALL)
					{
						//Regular elements
						strAcctName = actItm.strAcctName;

						//SOURCE:
						//		https://docs.microsoft.com/en-us/previous-versions/dotnet/netframework-4.0/ms256086(v=vs.100)
						//
						xNode = xRoot.SelectSingleNode(kgStrNodeNm_Account + "[not(@" + kgStrNodeAttNm_All + ") and @" +
							kgStrNodeAttNm_Name + "=" + escapeXPathValue(strAcctName) + "]");
					}
					else
					{
						//ALL elements
						xNode = xRoot.SelectSingleNode(kgStrNodeNm_Account + "[@" + kgStrNodeAttNm_All + "]");
						strAcctName = "";
					}

					if (xNode == null)
					{
						//Add such note
						XmlNode xNew = _xmlHdrs.CreateElement(kgStrNodeNm_Account);

						if (actItm.type != AccountItmType.ACCT_ITP_ALL)
						{
							XmlAttribute xAtt = _xmlHdrs.CreateAttribute(kgStrNodeAttNm_Name);
							xAtt.Value = strAcctName;
							xNew.Attributes.Append(xAtt);
						}
						else
						{
							XmlAttribute xAtt = _xmlHdrs.CreateAttribute(kgStrNodeAttNm_All);
							xNew.Attributes.Append(xAtt);
						}

						xRoot.AppendChild(xNew);

					}
				}


				//Remove nodes for accounts that we no longer have
				int nCntNodes = xRoot.ChildNodes.Count;
				bool bHaveAllAlready = false;

				//Go from the back forward
				for (int n = nCntNodes - 1; n >= 0; n--)
				{
					xNode = xRoot.ChildNodes[n];
					bool bDelNode = true;

					//Get node type
					XmlAttribute xAtt = xNode.Attributes[kgStrNodeAttNm_All];
					bool bAll = xAtt != null;

					if (!bAll)
					{
						//Non-All item
						xAtt = xNode.Attributes[kgStrNodeAttNm_Name];
						string strName = (xAtt != null) ? xAtt.Value : null;

						if (!string.IsNullOrEmpty(strName))
						{
							//See if we have such account in the list
							for (int i = 0; i < nCnt; i++)
							{
								AccountItem actItm = (AccountItem)comboBoxAccounts.Items[i];

								if (string.Compare(actItm.strAcctName, strName, false) == 0)
								{
									//We have such
									bDelNode = false;
									break;
								}
							}
						}
					}
					else
					{
						//ALL item -- can be only one
						if (!bHaveAllAlready)
						{
							bHaveAllAlready = true;
							bDelNode = false;
						}
					}

					if (bDelNode)
					{
						//Delete this node
						xRoot.RemoveChild(xNode);
					}
				}




				//Select first account
				if (comboBoxAccounts.Items.Count != 0)
				{
					comboBoxAccounts.SelectedIndex = 0;
				}

				//Reload the items in the list
				reloadHeadersForAccont();


				//MessageBox.Show("cnt=" + nCntNodes + "\n\n" + _xmlHdrs.OuterXml);
			}
			catch (Exception ex)
			{
				ThisAddIn.LogDiagnosticSpecError(1008, ex);
				MessageBox.Show("ERROR: Failed to load Outlook accounts: " + ex.ToString(), GlobalDefs.GlobDefs.gkstrAppNameFull, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally
			{
				//Set global flags back
				_bAllowCatchCmbAcctsChanges = true;
			}
		}


		public string escapeXPathValue(string value)
		{
			//RETURN:
			//		= Escaped XPath value

			//SOURCE:
			//		https://stackoverflow.com/a/1352556/843732
			//

			// if the value contains only single or double quotes, construct
			// an XPath literal
			if (!value.Contains("\""))
			{
				return "\"" + value + "\"";
			}
			if (!value.Contains("'"))
			{
				return "'" + value + "'";
			}

			// if the value contains both single and double quotes, construct an
			// expression that concatenates all non-double-quote substrings with
			// the quotes, e.g.:
			//
			//    concat("foo", '"', "bar")
			StringBuilder sb = new StringBuilder();
			sb.Append("concat(");
			string[] substrings = value.Split('\"');
			for (int i = 0; i < substrings.Length; i++)
			{
				bool needComma = (i > 0);
				if (substrings[i] != "")
				{
					if (i > 0)
					{
						sb.Append(", ");
					}
					sb.Append("\"");
					sb.Append(substrings[i]);
					sb.Append("\"");
					needComma = true;
				}
				if (i < substrings.Length - 1)
				{
					if (needComma)
					{
						sb.Append(", ");
					}
					sb.Append("'\"'");
				}

			}
			sb.Append(")");
			return sb.ToString();
		}


        private void btnApply_Click(object sender, EventArgs e)
        {
			//Start saving data
			if(saveData())
			{
				//Close this form with OK result
				this.DialogResult = DialogResult.OK;
				this.Close();
			}
			else
				ThisAddIn.LogDiagnosticSpecError(1009);
		}

		public bool saveData()
		{
			//Save data in this add-in to persistent storage
			//RETURN:
			//		= false if failed (will show message box error internally)

			//Start saving data
			string strErrDesc;
			bool bRes = false;

			//Set cursor
			Cursor.Current = Cursors.WaitCursor;

			try
			{
				//Convert to XML
				string strXml = getAccountHeadersToXMLString(out strErrDesc);
				if (strXml != null)
				{
					//Save data now
					bRes = ThisAddIn.saveHeaders(strXml, out strErrDesc);
					if(!bRes)
						ThisAddIn.LogDiagnosticSpecError(1012, strErrDesc);
				}
				else
					ThisAddIn.LogDiagnosticSpecError(1011, strErrDesc);
			}
			finally
			{
				//Restore cursor
				Cursor.Current = Cursors.Default;
			}

			if (bRes)
			{
				//And reset "dirty" state
				setDirtyState(false);
			}
			else
			{
				//Show error
				ThisAddIn.LogDiagnosticSpecError(1010, strErrDesc);
				MessageBox.Show("ERROR: Failed to save data.\n\n" + strErrDesc, GlobalDefs.GlobDefs.gkstrAppNameFull, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}

			return bRes;
		}

		private static void readNodeAttIntoTriState(XmlNode xmlNode, string strAttName, ref TriState outState)
		{
			XmlAttribute xAtt = xmlNode.Attributes[strAttName];
			if (xAtt != null)
			{
				string strV = xAtt.Value;

				if (string.Compare(strV, "1", true) == 0 ||
					string.Compare(strV, "true", true) == 0 ||
					string.Compare(strV, "on", true) == 0)
				{
					outState = TriState.TS_True;
				}
				else if (string.Compare(strV, "0", true) == 0 ||
					string.Compare(strV, "false", true) == 0 ||
					string.Compare(strV, "off", true) == 0)
				{
					outState = TriState.TS_False;
				}
				else if(string.Compare(strV, "2", true) == 0)
				{
					outState = TriState.TS_Inherit;
				}
			}
		}

		private static void readNodeAttIntoStr(XmlNode xmlNode, string strAttName, bool bTrimStr, ref string outStr)
		{
			XmlAttribute xAtt = xmlNode.Attributes[strAttName];
			if (xAtt != null)
			{
				string strV = xAtt.Value;

				if (bTrimStr)
					strV = strV.Trim();

				outStr = strV;
			}
		}

		public static OutboundEmailInfo getOutboundEmlInfoFromNode(XmlNode xmlNode)
		{
			//RETURN:
			//		= Email formatting info from 'xmlNode'
			OutboundEmailInfo efi = new OutboundEmailInfo();

			readNodeAttIntoTriState(xmlNode, kgStrNodeAttNm_OverwriteMailClientNm, ref efi.useOverwriteMailClient);
			readNodeAttIntoStr(xmlNode, kgStrNodeAttNm_MailClientNm, true, ref efi.strMailClientName);

			readNodeAttIntoTriState(xmlNode, kgStrNodeAttNm_SuppressRcvdHdr, ref efi.useSuppressRcvdHeader);
			readNodeAttIntoStr(xmlNode, kgStrNodeAttNm_RcvdHdr, true, ref efi.strRcvdHeader);


			XmlAttribute xAttSendAs = xmlNode.Attributes[kgStrNodeAttNm_SendAs];
			if (xAttSendAs != null)
			{
				//Convert the value
				int v;
				if (int.TryParse(xAttSendAs.Value, out v))
				{
					if (v >= 0 && v < (int)EmailFmt.EML_FMT_Count)
					{
						efi.sendEmlAs = (EmailFmt)v;
					}
				}
			}

			readNodeAttIntoTriState(xmlNode, kgStrNodeAttNm_SetCSSStyles, ref efi.useCssStyles);
			readNodeAttIntoStr(xmlNode, kgStrNodeAttNm_CSSStyles, true, ref efi.strCssStyles);

			readNodeAttIntoTriState(xmlNode, kgStrNodeAttNm_RemTrackPixel, ref efi.useRemoveTrackingPixel);
			readNodeAttIntoStr(xmlNode, kgStrNodeAttNm_TrackDoms, true, ref efi.strTrackingDomains);


			return efi;
		}

		public static ItemInfo getHeaderInfoFromNode(XmlNode xmlNode)
		{
			//RETURN:
			//		= Header item info from 'xmlNode', or
			//		= null if error

			XmlAttribute xAttName = xmlNode.Attributes[kgStrNodeAttNm_Name];
			if (xAttName == null)
			{
				ThisAddIn.LogDiagnosticSpecError(1013);
				return null;
			}

			ItemInfo ii = new ItemInfo();

			ii.strName = xAttName.Value;
			ii.strVal = xmlNode.InnerText;

			//Enabled?
			XmlAttribute xAttEnbld = xmlNode.Attributes[kgStrNodeAttNm_Enabled];
			if (xAttEnbld != null)
			{
				string str = xAttEnbld.Value.ToLower();
				if (str == "0" || str == "false")
				{
					ii.bEnabled = false;
				}
			}

			return ii;
		}

		private bool setHeaderListItem(int idxItem, bool bAllItem, XmlNode xmlNode)
		{
			//'bAllItem' = true if this is an ALL item
			//'idxItem' = index in 'listHdrs' of the list item to set, or -1 to add new item to the bottom
			//'xmlNode' = where to take the data from
			//RETURN:
			//		= true if success
			bool bRes = false;

			if (_acctItem == null)
			{
				ThisAddIn.LogDiagnosticSpecError(1014);
				return false;
			}

			//Remember old value
			bool bOldVal = _bAllowCatchListHdrsChanges;

			//Prevent changes to the list, while we're setting things
			_bAllowCatchListHdrsChanges = false;

			ItemInfo ii = getHeaderInfoFromNode(xmlNode);
			if (ii != null)
			{
				bRes = true;
			}
			else
				ii = new ItemInfo();

			if (idxItem >= 0 && idxItem < listHdrs.Items.Count)
			{
				//Set existing
				listHdrs.Items[idxItem].SubItems[1].Text = ii.strName;
				listHdrs.Items[idxItem].SubItems[2].Text = ii.strVal;
			}
			else
			{
				//Add new
				ListViewItem lvi = new ListViewItem(new[] { bAllItem && _acctItem.type != AccountItmType.ACCT_ITP_ALL ? "ALL" : "", ii.strName, ii.strVal });

				lvi.Checked = ii.bEnabled;
				lvi.Tag = xmlNode;

				if(bAllItem)
				{
					//If we're not in the ALL selection
					if (_acctItem.type != AccountItmType.ACCT_ITP_ALL)
					{
						lvi.ForeColor = SystemColors.GrayText;
					}
				}

				listHdrs.Items.Add(lvi);
			}

			//Restore
			_bAllowCatchListHdrsChanges = bOldVal;

			return bRes;
		}

		private void addNewHeaderItem()
		{
			//Show a dialog to add new header item

			//Do we have an account node?
			if (_xmlHdrs == null || _xmlNodeSelAcct == null || _acctItem == null)
			{
				ThisAddIn.LogDiagnosticSpecError(1015);
				SystemSounds.Exclamation.Play();
				return;
			}

			//Add new item
			FormEditHeader feh = new FormEditHeader();
			feh.hdrType.xmlDoc = _xmlHdrs;
			feh.hdrType.xmlParNode = _xmlNodeSelAcct;

			if (feh.ShowDialog() == DialogResult.OK)
			{
				//Add new value
				setHeaderListItem(-1, _acctItem.type == AccountItmType.ACCT_ITP_ALL, feh.hdrType.xmlNode);

				//Select last item
				int nCntAll = listHdrs.Items.Count;
				if (nCntAll > 0)
				{
					int nIdx = nCntAll - 1;
					listHdrs.SelectedIndices.Clear();
					listHdrs.SelectedIndices.Add(nIdx);
					listHdrs.EnsureVisible(nIdx);
				}

				//Update UI
				updateUIControls();

				//Mark data as "dirty"
				setDirty("1");
			}

		}

        private void btnAddHdr_Click(object sender, EventArgs e)
        {
			addNewHeaderItem();
		}

		private HeaderItem getCurrentHeaderItemForEdit()
		{
			//RETURN:
			//		= Header item info for editing, or
			//		= null if not available
			HeaderItem hdrItm = getCurrentHeaderItem();
			if (hdrItm != null)
			{
				bool bAllowToEdit = false;
				if (_acctItem.type == AccountItmType.ACCT_ITP_ALL)
				{
					bAllowToEdit = true;
				}
				else
				{
					if (!hdrItm.bAllItem)
						bAllowToEdit = true;
				}

				if (bAllowToEdit)
				{
					return hdrItm;
				}
			}

			return null;
		}

		private void editSelectedHeaderItem(bool bAfterDoubleClick)
		{
			//Begin editing the currently selected header item
			if (_acctItem == null)
			{
				ThisAddIn.LogDiagnosticSpecError(1016);
				SystemSounds.Exclamation.Play();
				return;
			}

			HeaderItem hdrItm = getCurrentHeaderItemForEdit();
			if (hdrItm != null)
			{
				if(bAfterDoubleClick)
				{
					//This is needed to compensate for annoying Windows "feature" that toggle the checkbox upon a double-click
					hdrItm.lvItem.Checked = hdrItm.lvItem.Checked ? false : true;
				}

				//Show dialog to edit item
				FormEditHeader feh = new FormEditHeader();

				feh.hdrType.xmlDoc = _xmlHdrs;
				feh.hdrType.xmlParNode = _xmlNodeSelAcct;
				feh.hdrType.xmlNode = hdrItm.xmlNode;

				if (feh.ShowDialog() == DialogResult.OK)
				{
					//Update the list item
					setHeaderListItem(hdrItm.lvItem.Index, false, feh.hdrType.xmlNode);

					//Mark data as "dirty"
					setDirty("2");
				}
			}
		}

        private void listHdrs_DoubleClick(object sender, EventArgs e)
        {
            //Item double-clicked
			editSelectedHeaderItem(true);
        }

		private void aboutThisAddinToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//About this Add-in
			FormAbout fa = new FormAbout();
			fa.ShowDialog();
		}

		private string getAccountHeadersToXMLString(out string strOutErrDesc)
		{
			//'strOutErrDesc' = receive error description, if any
			//RETURN:
			//		= XML document string for current account headers, or
			//		= null if error
			strOutErrDesc = "";
			string strXml = null;

			if (_xmlHdrs != null)
			{
				try
				{
					using (StringWriter sW = new StringWriter())
					{
						using (XmlTextWriter xW = new XmlTextWriter(sW))
						{
							_xmlHdrs.WriteTo(xW);

							strXml = sW.ToString();
						}
					}
				}
				catch (Exception ex)
				{
					ThisAddIn.LogDiagnosticSpecError(1017, ex);
					strOutErrDesc = "Failed to convert XML to string: " + ex.ToString();
					strXml = null;
				}
			}
			else
				strOutErrDesc = "No XML headers object";

			return strXml;
		}



		private void exportToFileToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Export to a file
			string strErrDesc;
			string strXml = getAccountHeadersToXMLString(out strErrDesc);
			if (strXml != null)
			{
				//Show dialog to pick file for saving
				SaveFileDialog sfd = new SaveFileDialog();
				sfd.Title = "Select Where To Save Exported Data";
				sfd.Filter = GlobalDefs.GlobDefs.gkstrAppNameFull + " Exports|*." + GlobalDefs.GlobDefs.gkstrExportFileExt + "|All Files|*.*";
				sfd.DefaultExt = "." + GlobalDefs.GlobDefs.gkstrExportFileExt;

				//Default file name
				DateTime dtNow = DateTime.Now;
				sfd.FileName = GlobalDefs.GlobDefs.gkstrAppNameFull + " Data Export, " + dtNow.Year.ToString("D4") + "-" + dtNow.Month.ToString("D2") + "-" + dtNow.Day.ToString("D2");

				if (sfd.ShowDialog() == DialogResult.OK)
				{
					//Save now
					Cursor.Current = Cursors.WaitCursor;

					try
					{
						string strFilePath = sfd.FileName;

						//Set encoding (with a BOM)
						Encoding enc = new UTF8Encoding(true);
						using (TextWriter tW = new StreamWriter(strFilePath, false, enc))
						{
							tW.Write(strXml);
						}
					}
					finally
					{
						//Restore cursor
						Cursor.Current = Cursors.Default;
					}
				}
			}

			//See if any errors?
			if (!string.IsNullOrEmpty(strErrDesc))
			{
				//Error
				ThisAddIn.LogDiagnosticSpecError(1018, strErrDesc);
				MessageBox.Show("ERROR: Failed to export data.\n\n" + strErrDesc, GlobalDefs.GlobDefs.gkstrAppNameFull, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}


		private void importFromFileToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Import data from an outside file
			string strErrDesc;

			//Show open file dialog
			OpenFileDialog ofd = new OpenFileDialog();
			ofd.Title = "Select Previously Saved File to Import";
			ofd.Filter = GlobalDefs.GlobDefs.gkstrAppNameFull + " Exports|*." + GlobalDefs.GlobDefs.gkstrExportFileExt + "|All Files|*.*";
			ofd.Multiselect = false;

			if (ofd.ShowDialog() == DialogResult.OK)
			{
				//Save now
				Cursor.Current = Cursors.WaitCursor;

				string strFilePath = ofd.FileName;

				try
				{
					//Read data from the file
					string strXML = System.IO.File.ReadAllText(strFilePath);

					//And validate it while converting it to XML
					ImportedDoc impDoc = verifyXmlDataForImporting(strXML, out strErrDesc);
					if (impDoc != null)
					{
						//Validated all good, ask user if they want to do this
						if (MessageBox.Show("WARNING: You are about to import new headers from the following file:\n\n" + strFilePath +
							"\n\n" +
							"This will override all current headers in this add-in.\n\n" +
							"Do you want to continue?",
							GlobalDefs.GlobDefs.gkstrAppNameFull, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
						{
							//Can use the data now
							_xmlHdrs = impDoc.xmlDoc;

							//And reload everything in the window
							reloadAllAcoounts();

							//Mark data as "dirty"
							setDirty("3");
						}
					}
					else
					{
						//Error
						ThisAddIn.LogDiagnosticSpecError(1019, strErrDesc);
						MessageBox.Show("ERROR: Failed to validate the following file:\n\n" + strFilePath + "\n\n" + strErrDesc,
							GlobalDefs.GlobDefs.gkstrAppNameFull, MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
				catch (Exception ex)
				{
					//Exception
					ThisAddIn.LogDiagnosticSpecError(1020, ex);
					MessageBox.Show("ERROR: Failed to read file:\n\n" + strFilePath + "\n\n" + ex.ToString(),
						GlobalDefs.GlobDefs.gkstrAppNameFull, MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				finally
				{
					//Restore cursor
					Cursor.Current = Cursors.Default;
				}
			}
			

		}


		private void deleteAllHeadersToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Delete all header data
			if (MessageBox.Show("WARNING: You are about to delete all headers for all accounts in this add-in.\n\n" +
				"Do you want to continue?",
				GlobalDefs.GlobDefs.gkstrAppNameFull, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
			{
				//Go ahead with removal
				Cursor.Current = Cursors.WaitCursor;

				try
				{
					//First clear XML object
					_xmlHdrs.RemoveAll();
					_xmlHdrs = null;

					_xmlHdrs = new XmlDocument();

					//And reload the UI
					reloadAllAcoounts();

					//Mark data as "dirty"
					setDirty("4");
				}
				finally
				{
					//Restore cursor
					Cursor.Current = Cursors.Default;
				}
			}

		}

		private void comboBoxAccounts_SelectedIndexChanged(object sender, EventArgs e)
		{
			//Selected account changed
			if (!_bAllowCatchCmbAcctsChanges)
				return;

			reloadHeadersForAccont();
		}


		private XmlNode findXmlNodeForAccount(AccountItem actItm)
		{
			//RETURN:
			//		= XML node for 'actItm'
			//		= null if none
			if (_xmlHdrs == null)
			{
				ThisAddIn.LogDiagnosticSpecError(1021);
				return null;
			}

			//Get root
			XmlElement xRoot = _xmlHdrs.DocumentElement;

			//Find this XML node
			if (actItm.type != AccountItmType.ACCT_ITP_ALL)
			{
				//Regulat elements

				//SOURCE:
				//		https://docs.microsoft.com/en-us/previous-versions/dotnet/netframework-4.0/ms256086(v=vs.100)
				//
				return xRoot.SelectSingleNode(kgStrNodeNm_Account + "[not(@" + kgStrNodeAttNm_All + ") and @" +
					kgStrNodeAttNm_Name + "=" + escapeXPathValue(actItm.strAcctName) + "]");
			}
			else
			{
				//ALL elements
				return xRoot.SelectSingleNode(kgStrNodeNm_Account + "[@" + kgStrNodeAttNm_All + "]");
			}
		}

		private void reloadHeadersForAccont()
		{
			//Reload all headers for the account

			//Set flags
			_bAllowCatchListHdrsChanges = false;

			try
			{
				//Remove selected node
				_xmlNodeSelAcct = null;
				_acctItem = null;

				//Remove everything from the list
				listHdrs.Items.Clear();

				//Clear all formatting options
				comboBoxSendFmt.Items.Clear();
				checkBoxRemTrPxl.Checked = false;

				string strAcctName = "";
				bool bAllAcct = false;
				OutboundEmailInfo efi = null;

				//Which item is selected
				int nSelIdx = comboBoxAccounts.SelectedIndex;
				if (nSelIdx >= 0 && nSelIdx < comboBoxAccounts.Items.Count)
				{
					//Selected account
					AccountItem actItm = (AccountItem)comboBoxAccounts.Items[nSelIdx];
					strAcctName = actItm.ToString();

					//Find this XML node
					_xmlNodeSelAcct = findXmlNodeForAccount(actItm);

					if (_xmlNodeSelAcct != null)
					{
						//Load items for this node
						_acctItem = actItm;

						//Only if we're not in the ALL account
						bAllAcct = _acctItem.type == AccountItmType.ACCT_ITP_ALL;
						if (!bAllAcct)
						{
							//First load ALL headers
							XmlElement xRoot = _xmlHdrs.DocumentElement;

							XmlNode xAll = xRoot.SelectSingleNode(kgStrNodeNm_Account + "[@" + kgStrNodeAttNm_All + "]");
							if (xAll != null)
							{
								foreach (XmlNode xNode in xAll.ChildNodes)
								{
									setHeaderListItem(-1, true, xNode);
								}
							}
						}

						//Then load items from this node
						foreach (XmlNode xNode in _xmlNodeSelAcct)
						{
							setHeaderListItem(-1, false, xNode);
						}


						//Fill in "Send emails in format:" combo
						efi = getOutboundEmlInfoFromNode(_xmlNodeSelAcct);

						EmlFmt_Item[] efiItms = new EmlFmt_Item[]{
							new EmlFmt_Item() { emlFmt = EmailFmt.EML_FMT_None, strName = bAllAcct ? "" : "<Inherit>"},
							new EmlFmt_Item() { emlFmt = EmailFmt.EML_FMT_Dont_Change, strName = "Don't Change"},
							new EmlFmt_Item() { emlFmt = EmailFmt.EML_FMT_SendAsText, strName = "Plain text"},
							new EmlFmt_Item() { emlFmt = EmailFmt.EML_FMT_SendAsHTML, strName = "HTML"},
						};

						foreach(EmlFmt_Item ei in efiItms)
						{
							comboBoxSendFmt.Items.Add(ei);

							if (ei.emlFmt == efi.sendEmlAs)
							{
								//Select it
								comboBoxSendFmt.SelectedItem = ei;
							}
						}

					}
					else
					{
						//Error
						string strDbgInf = actItm.toDebugString();
						ThisAddIn.LogDiagnosticSpecError(1022, strDbgInf);
						MessageBox.Show("ERROR: Failed to find node: " + strDbgInf, GlobalDefs.GlobDefs.gkstrAppNameFull, MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}


				//Set group-box name
				groupBoxMain.Text = "Custom headers for:  " + strAcctName;

				//Set checkboxes to tri-state if not ALL account
				checkBoxOverwriteMailClientNm.ThreeState = bAllAcct ? false : true;
				checkBoxSuppressRcvd.ThreeState = bAllAcct ? false : true;
				checkBoxCssStyle.ThreeState = bAllAcct ? false : true;
				checkBoxRemTrPxl.ThreeState = bAllAcct ? false : true;

				//Set checkboxes
				setThreeStateCheckFromTriState(checkBoxOverwriteMailClientNm, efi != null ? efi.useOverwriteMailClient : TriState.TS_Inherit, !bAllAcct);
				setThreeStateCheckFromTriState(checkBoxSuppressRcvd, efi != null ? efi.useSuppressRcvdHeader : TriState.TS_Inherit, !bAllAcct);
				setThreeStateCheckFromTriState(checkBoxCssStyle, efi != null ? efi.useCssStyles : TriState.TS_Inherit, !bAllAcct);
				setThreeStateCheckFromTriState(checkBoxRemTrPxl, efi != null ? efi.useRemoveTrackingPixel : TriState.TS_Inherit, !bAllAcct);

				//Set text boxes
				textBoxMailClientNm.Text = efi != null ? efi.strMailClientName.Trim() : "";
				textBoxSuppressRcvd.Text = efi != null ? efi.strRcvdHeader.Trim() : "";
				textBoxCssStyle.Text = efi != null ? efi.strCssStyles.Trim() : "";
				textBoxTrckDomains.Text = efi != null ? efi.strTrackingDomains.Trim() : "";


				//Update controls
				updateOverwriteMailClientNameCtrls();
				updateSuppressReceivedHaderCtrls();
				updateEmailFmtCtrls();
				updateTrackingPixelsCtrls();
			}
			finally
			{
				//Reset flags
				_bAllowCatchListHdrsChanges = true;
			}
		}

		private HeaderItem getCurrentHeaderItem()
		{
			//RETURN:
			//		= Currently selected header item, or
			//		= null if none
			if (listHdrs.SelectedIndices.Count > 0)
			{
				return getHeaderItem(listHdrs.SelectedIndices[0]);
			}

			return null;
		}

		private HeaderItem getHeaderItem(int idxItm)
		{
			//RETURN:
			//		= Header item info for index 'idxItm', or
			//		= null if none
			if (_acctItem != null)
			{
				//Get selected item
				if (idxItm >= 0 && idxItm < listHdrs.Items.Count)
				{
					ListViewItem itm = listHdrs.Items[idxItm];
					if (itm != null)
					{
						if(itm.Tag != null && itm.Tag is XmlNode)
						{
							HeaderItem hi = new HeaderItem();

							hi.xmlNode = (XmlNode)itm.Tag;
							hi.lvItem = itm;
							
							if (_acctItem.type != AccountItmType.ACCT_ITP_ALL)
							{
								hi.bAllItem = string.IsNullOrEmpty(itm.SubItems[0].Text) ? false : true;
							}
							else
							{
								hi.bAllItem = false;
							}

							return hi;
						}
					}
				}
			}

			return null;
		}


		private void listHdrs_ItemCheck(object sender, ItemCheckEventArgs e)
		{
			//When the list item was checked or unchecked
			if (!_bAllowCatchListHdrsChanges)
				return;

			bool bKeepSame = true;

			//New check
			bool bChecked = e.NewValue == CheckState.Checked;

			//Get selected item
			HeaderItem hdrItm = getHeaderItem(e.Index);
			if (hdrItm != null)
			{
				if (!hdrItm.bAllItem)
				{
					//Estimate location of this checkbox on the screen
					//INFO: We need to do this because for some unknown to me reason list-view wants to check
					//      controls for any ungodly reasons, like selection, etc.
					//      I have no time to root out these old Microsoft bugs....
					Point pnt = listHdrs.PointToClient(Cursor.Position);
					Rectangle rcSubitem0 = listHdrs.Items[e.Index].SubItems[0].Bounds;		//It is set to width of the entire row
					Rectangle rcSubitem1 = listHdrs.Items[e.Index].SubItems[1].Bounds;

					//Calculate size of column with the checkbox
					Rectangle rcChkbxColumn = new Rectangle(rcSubitem0.X, rcSubitem0.Y, rcSubitem1.X - rcSubitem0.X, rcSubitem0.Width);

					//Is the mouse in that area?
					if (rcChkbxColumn.Contains(pnt))
					{
						bKeepSame = false;

						//Set check
						setHeaderItemEnabled(hdrItm.xmlNode, bChecked);

						//Mark data as "dirty"
						setDirty("5");
					}
				}
			}

			if (bKeepSame)
			{
				//Don't allow to change
				e.NewValue = e.CurrentValue;
			}
		}

		private void setHeaderItemEnabled(XmlNode xmlNode, bool bChecked)
		{
			//Set check
			XmlAttribute xAtt = xmlNode.Attributes[kgStrNodeAttNm_Enabled];
			if (xAtt == null)
			{
				xAtt = _xmlHdrs.CreateAttribute(kgStrNodeAttNm_Enabled);
				xAtt.Value = bChecked ? "1" : "0";

				xmlNode.Attributes.Append(xAtt);
			}
			else
			{
				xAtt.Value = bChecked ? "1" : "0";
			}
		}

		private void listHdrs_MouseClick(object sender, MouseEventArgs e)
		{
			//Pick mouse right-clicks
			if (e.Button == MouseButtons.Right)
			{
				contextMenuHdrsLst.Show(Cursor.Position);
			}
		}

		private void contextMenuHdrsLst_Opening(object sender, CancelEventArgs e)
		{
			//This is called when the main config window's context menu is about to be shown

			//Enable/disable items
			HeaderItem hdrItm = getCurrentHeaderItemForEdit();
			bool bAllowRemove = removeSelectedHeaderItems(true);
			bool bAllowMoveUp = moveSelectedHeaderItems(true, true);
			bool bAllowMoveDown = moveSelectedHeaderItems(false, true);

			editToolStripMenuItm.Enabled = hdrItm != null;
			removeToolStripMenuItem.Enabled = bAllowRemove;
			moveUpToolStripMenuItem.Enabled = bAllowMoveUp;
			moveDownToolStripMenuItem.Enabled = bAllowMoveDown;

			//How many items are selected in the list
			bool bAnySelection = listHdrs.SelectedItems.Count > 0;
			checkSelectedToolStripMenuItem.Enabled = bAnySelection;
			uncheckSelectedToolStripMenuItem.Enabled = bAnySelection;
		}

		private void menuStrip1_Click(object sender, EventArgs e)
		{
			//Try to update items for the main menu
			updateUIControls();
		}

		private void updateUIControls()
		{
			//Update UI controls and menus
			HeaderItem hdrItm = getCurrentHeaderItemForEdit();
			bool bAllowRemove = removeSelectedHeaderItems(true);
			bool bAllowMoveUp = moveSelectedHeaderItems(true, true);
			bool bAllowMoveDown = moveSelectedHeaderItems(false, true);

			editHeaderToolStripMenuItem.Enabled = hdrItm != null;
			deleteHeaderToolStripMenuItem.Enabled = bAllowRemove;
			moveHeadersUpToolStripMenuItem.Enabled = bAllowMoveUp;
			moveHeadersDownToolStripMenuItem.Enabled = bAllowMoveDown;

			//Remove button
			btnDelHdr.Enabled = bAllowRemove;
		}


		private void editToolStripMenuItm_Click(object sender, EventArgs e)
		{
			//Edit this item
			editSelectedHeaderItem(false);
		}

		private void addNewToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Add new item
			addNewHeaderItem();
		}

		private bool checkUncheckAllHeaderItems(bool bCheck, bool bOnlySelected)
		{
			//Check or uncheck all (available) header items in the list
			//'bOnlySelected' = true to do this only to selected items, false - to do it to all items
			//RETURN:
			//		= true if success
			if (_acctItem == null)
			{
				ThisAddIn.LogDiagnosticSpecError(1023);
				return false;
			}

			//Prevent catching changes
			_bAllowCatchListHdrsChanges = false;

			int nCnt = listHdrs.Items.Count;
			for (int i = 0; i < nCnt; i++)
			{
				if (bOnlySelected)
				{
					//Check if item is selected
					if (!listHdrs.Items[i].Selected)
						continue;
				}

				HeaderItem hdrItm = getHeaderItem(i);
				if (hdrItm != null)
				{
					bool bAllow = false;
					if (_acctItem.type != AccountItmType.ACCT_ITP_ALL)
					{
						if (!hdrItm.bAllItem)
							bAllow = true;
					}
					else
					{
						bAllow = true;
					}

					if (bAllow)
					{
						//Can set it
						setHeaderItemEnabled(hdrItm.xmlNode, bCheck);

						//Update list
						hdrItm.lvItem.Checked = bCheck;
					}
				}
			}

			//Restore
			_bAllowCatchListHdrsChanges = true;

			//Mark data as "dirty"
			setDirty("6");

			return true;
		}

		private void checkAllToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Check all items
			if (!checkUncheckAllHeaderItems(true, false))
			{
				//Failed
				ThisAddIn.LogDiagnosticSpecError(1024);
				SystemSounds.Exclamation.Play();
			}
		}

		private void uncheckAllToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Un-check all items
			if (!checkUncheckAllHeaderItems(false, false))
			{
				//Failed
				ThisAddIn.LogDiagnosticSpecError(1025);
				SystemSounds.Exclamation.Play();
			}
		}

		private bool removeSelectedHeaderItems(bool bTestOnly)
		{
			//Remove selected header item(s) after showing a warning to the user
			//'bTestOnly' = true to test removal, false = actually remove
			//RETURN:
			//		= true if no errors
			if (_acctItem == null)
			{
				ThisAddIn.LogDiagnosticSpecError(1026);
				return false;
			}

			//Collect items that can be removed
			List<HeaderItem> arrDelItms = new List<HeaderItem>();

			int nCnt = listHdrs.SelectedIndices.Count;
			for (int i = 0; i < nCnt; i++)
			{
				int idxSel = listHdrs.SelectedIndices[i];
				HeaderItem hi = getHeaderItem(idxSel);
				if (hi != null)
				{
					if (!hi.bAllItem)
					{
						//Can delete this on
						arrDelItms.Add(hi);
					}
				}
			}

			bool bRes = false;

			//Did we get any?
			int nCntDel = arrDelItms.Count;
			if (nCntDel > 0)
			{
				//Assime success
				bRes = true;

				if (!bTestOnly)
				{
					//Make a warning
					string strMsgWarn;
					if (nCntDel == 1)
					{
						strMsgWarn = "Do you want to remove the following header?\n\n" + arrDelItms[0].lvItem.SubItems[1].Text + " = " + arrDelItms[0].lvItem.SubItems[2].Text;
					}
					else
					{
						strMsgWarn = "Do you want to remove " + nCntDel.ToString() + " selected headers?";
					}

					//Show user warning
					if (MessageBox.Show(strMsgWarn,
						GlobalDefs.GlobDefs.gkstrAppNameFull, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
					{
						//Can remove now

						//Set cursor
						Cursor.Current = Cursors.WaitCursor;

						try
						{
							try
							{
								//Go through all items
								for (int r = 0; r < nCntDel; r++)
								{
									HeaderItem hdtItm = arrDelItms[r];

									//Remove XML node
									hdtItm.xmlNode.ParentNode.RemoveChild(hdtItm.xmlNode);

									//Remove UI element
									listHdrs.Items.Remove(hdtItm.lvItem);
								}

							}
							catch (Exception ex)
							{
								//Failed
								ThisAddIn.LogDiagnosticSpecError(1027, ex);
								MessageBox.Show("ERROR: Failed to remove header item(s): " + ex.ToString(), GlobalDefs.GlobDefs.gkstrAppNameFull, MessageBoxButtons.OK, MessageBoxIcon.Error);
								bRes = false;
							}
						}
						finally
						{
							//Restore cursor
							Cursor.Current = Cursors.Default;

							//Update controls
							updateUIControls();

							//Mark data as "dirty"
							setDirty("7");
						}
					}
				}
			}

			return bRes;
		}

		private void btnDelHdr_Click(object sender, EventArgs e)
		{
			//Remove selected items
			if (!removeSelectedHeaderItems(false))
			{
				//Failed
				ThisAddIn.LogDiagnosticSpecError(1028);
				SystemSounds.Exclamation.Play();
			}
		}

		private void removeToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Remove selected items
			if (!removeSelectedHeaderItems(false))
			{
				//Failed
				ThisAddIn.LogDiagnosticSpecError(1029);
				SystemSounds.Exclamation.Play();
			}
		}

		private void moveUpToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Move selected items up
			if (!moveSelectedHeaderItems(true, false))
			{
				ThisAddIn.LogDiagnosticSpecError(1030);
				SystemSounds.Exclamation.Play();
			}
		}

		private void moveDownToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Move selected items down
			if (!moveSelectedHeaderItems(false, false))
			{
				ThisAddIn.LogDiagnosticSpecError(1031);
				SystemSounds.Exclamation.Play();
			}
		}

		private void editHeaderToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Edit selected item
			editSelectedHeaderItem(false);
		}

		private void addNewHeaderToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Add new header
			addNewHeaderItem();
		}

		private void moveHeadersUpToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//move items up
			if (!moveSelectedHeaderItems(true, false))
			{
				ThisAddIn.LogDiagnosticSpecError(1032);
				SystemSounds.Exclamation.Play();
			}

		}

		private void moveHeadersDownToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Move items down
			if (!moveSelectedHeaderItems(false, false))
			{
				ThisAddIn.LogDiagnosticSpecError(1033);
				SystemSounds.Exclamation.Play();
			}

		}

		private void deleteHeaderToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Delete selected header items
			if (!removeSelectedHeaderItems(false))
			{
				//Failed
				ThisAddIn.LogDiagnosticSpecError(1034);
				SystemSounds.Exclamation.Play();
			}
		}



		private bool moveSelectedHeaderItems(bool bMoveUp, bool bTestOnly)
		{
			//Move selected header items up or down
			//'bMoveUp' = true to move up, false - to move down
			//'bTestOnly' = true to test if we can move, false = to actually move
			//RETURN:
			//		= true if no error
			if (_acctItem == null)
			{
				ThisAddIn.LogDiagnosticSpecError(1035);
				return false;
			}

			bool bRes = false;

			//Collect items that can be moved (sort them in ascending order by the index in the list)
			List<HeaderItem> arrMovItms = new List<HeaderItem>();

			int nCntSel = listHdrs.SelectedIndices.Count;
			if (nCntSel > 0)
			{
				//Assume success
				bRes = true;

				//Go thru all selected items
				for (int i = 0; i < nCntSel; i++)
				{
					int idxSel = listHdrs.SelectedIndices[i];
					HeaderItem hi = getHeaderItem(idxSel);
					if (hi == null || hi.bAllItem)
					{
						bRes = false;
						break;
					}

					//Can move this one
					int nThisIdx = hi.lvItem.Index;

					//Add in ascending order by index
					int rAt = -1;
					int nCntRem = arrMovItms.Count;
					for (int r = 0; r < nCntRem; r++)
					{
						int idx = arrMovItms[r].lvItem.Index;
						if (nThisIdx < idx)
						{
							//Add here
							rAt = r;
							break;
						}
					}

					if (rAt >= 0)
						arrMovItms.Insert(rAt, hi);
					else
						arrMovItms.Add(hi);
				}
			}

			if (bRes)
			{
				int nCntMov = arrMovItms.Count;
				if (nCntMov > 0)
				{
					if (!bTestOnly)
					{
						//Reset flag not to track changes in the list
						_bAllowCatchListHdrsChanges = false;
					}

					try
					{
						int nBegin, nEnd, nDir, nMovDir;
						if (bMoveUp)
						{
							nBegin = 0;
							nEnd = nCntMov - 1;
							nDir = 1;
							nMovDir = -1;
						}
						else
						{
							nBegin = nCntMov - 1;
							nEnd = 0;
							nDir = -1;
							nMovDir = 1;
						}

						int nCntLstItms = listHdrs.Items.Count;

						bool bFirstIter = true;

						for (int i = nBegin;; i += nDir, bFirstIter = false)
						{
							HeaderItem hdrItmF = arrMovItms[i];

							int idxLstF = hdrItmF.lvItem.Index;
							int idxLstT = idxLstF + nMovDir;

							HeaderItem hdrItmT = getHeaderItem(idxLstT);
							if (hdrItmT == null || hdrItmT.bAllItem)
							{
								//Can't do that
								bRes = false;
								break;
							}

							if (bFirstIter)
							{
								//Can we go ahead with the move?
								if (bTestOnly)
									break;
							}

							//Sanity check - Both nodes must have the same parent
							XmlNode xmlPar = hdrItmT.xmlNode.ParentNode;
							if (xmlPar != hdrItmF.xmlNode.ParentNode)
							{
								//Error
								ThisAddIn.LogDiagnosticSpecError(1036);
								bRes = false;
								break;
							}


							//Special move actions, depending on the direction
							if (bMoveUp)
							{
								//Swap two items in XML cache
								xmlPar.InsertAfter(hdrItmT.xmlNode, hdrItmF.xmlNode);

								//Swap two items in the UI
								listHdrs.Items.RemoveAt(idxLstT);
								listHdrs.Items.Insert(idxLstF, hdrItmT.lvItem);
							}
							else
							{
								//Swap two items in XML cache
								xmlPar.InsertAfter(hdrItmF.xmlNode, hdrItmT.xmlNode);

								//Swap two items in the UI
								listHdrs.Items.RemoveAt(idxLstF);
								listHdrs.Items.Insert(idxLstT, hdrItmF.lvItem);
							}

							if (bFirstIter)
							{
								//Set focus & make the item in the list visible
								listHdrs.Items[idxLstT].Focused = true;
								listHdrs.Items[idxLstT].EnsureVisible();
							}

							if (i == nEnd)
								break;
						}
					}
					catch (Exception ex)
					{
						//Failed
						ThisAddIn.LogDiagnosticSpecError(1037, ex);
						MessageBox.Show("ERROR: Failed to move header item(s): " + ex.ToString(), GlobalDefs.GlobDefs.gkstrAppNameFull, MessageBoxButtons.OK, MessageBoxIcon.Error);
						bRes = false;
					}
					finally
					{
						if (!bTestOnly)
						{
							//Set flag back
							_bAllowCatchListHdrsChanges = true;

							//Update UI
							updateUIControls();

							//Mark data as "dirty"
							setDirty("8");
						}
					}
				}
				else
					bRes = false;
			}

			return bRes;
		}


		Timer _tmrSelIdxChng = null;

		private void listHdrs_SelectedIndexChanged(object sender, EventArgs e)
		{
			//Selection in the list changed

			//INFO: Because there will be a lot of these messages, don't process them all here.
			//      Wait for 100ms in case another of these messages arrives...
			if (_tmrSelIdxChng == null)
			{
				_tmrSelIdxChng = new Timer();
				_tmrSelIdxChng.Interval = 100;
				_tmrSelIdxChng.Tick += new EventHandler(onTimer_SelectedIndexChanged);
			}
			else
				_tmrSelIdxChng.Stop();

			_tmrSelIdxChng.Start();
		}

		private void onTimer_SelectedIndexChanged(object sender, EventArgs e)
		{
			//Stop timer from running
			Timer tmr = (Timer)sender;
			tmr.Stop();

			//Update UI
			updateUIControls();
		}


		public static ImportedDoc verifyXmlDataForImporting(string strXML, out string strOutErrDesc)
		{
			//Verify that XML data in 'strXML' can be imported into this app
			//'strOutErrDesc' = receives error description, if any
			//RETURN:
			//		= Class for the data if it was accepted, or
			//		= null if error
			ImportedDoc resData = null;
			strOutErrDesc = "";

			if (!string.IsNullOrEmpty(strXML))
			{
				try
				{
					resData = new ImportedDoc();

					resData.xmlDoc = new XmlDocument();
					XmlDocument xDoc = resData.xmlDoc;

					//Load XML
					xDoc.LoadXml(strXML);

					bool bSuccess = false;

					//What is the name of the root node?
					if (xDoc.DocumentElement.Name == kgStrNodeNm_Root)
					{
						//Compare its app name
						if (xDoc.DocumentElement.Attributes[kgStrNodeAttNm_RootApp] != null &&
							xDoc.DocumentElement.Attributes[kgStrNodeAttNm_RootApp].Value == GlobalDefs.GlobDefs.gkstrAppNameFull)
						{
							//Get the version of the XML data
							string strVer = "";
							if (xDoc.DocumentElement.Attributes[kgStrNodeAttNm_RootVer] != null)
							{
								strVer = xDoc.DocumentElement.Attributes[kgStrNodeAttNm_RootVer].Value.Trim();

								//In this case we can process only v.1.0.0
								if (strVer == ThisAddIn.kstrDataFileVer)
								{
									//Remember file version
									resData.strVersion = strVer;


									//Go through all children nodes
									foreach (XmlNode xmlNode in xDoc.DocumentElement.ChildNodes)
									{
										string strNodeName = xmlNode.Name;
										if (strNodeName == kgStrNodeNm_Account)
										{
											//Check for the ALL node
											if (xmlNode.Attributes[kgStrNodeAttNm_All] != null)
											{
												//We have the ALL attribute -- assume success
												bSuccess = true;

												break;
											}

										}
									}
								}
								else
									strOutErrDesc = "The version of the data file ('" + strVer + "') is not supported by this version of the add-in. " +
										"Please download the latest version from " + ThisAddIn.kstrAppDownloadURL;
							}
						}
					}


					//Check results
					if (!bSuccess)
					{
						if (string.IsNullOrEmpty(strOutErrDesc))
						{
							//Validation of data didn't succeed
							strOutErrDesc = "Failed to validate XML data.";
						}

						//Free the doc
						resData = null;
					}
				}
				catch (Exception ex)
				{
					ThisAddIn.LogDiagnosticSpecError(1038, ex);
					strOutErrDesc = "Failed to parse: " + ex.ToString();
					resData = null;
				}
			}
			else
				strOutErrDesc = "No XML data found.";

			return resData;
		}

		private void saveToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Save data in this add-in
			saveData();
		}

		private void setDirty(string _debugFromWhere)
		{
			//Set this data as "dirty", or needing saving
			setDirtyState(true);

			//MessageBox.Show(_debugFromWhere);
		}

		private void setDirtyState(bool bNewDirtyState)
		{
			//Set this data as "dirty"
			//'bNewDirtyState' = true if it needs saving, false - if data is saved
			_gbDirty = bNewDirtyState;

			//Update title of the window
			this.Text = "Customize Outbound Mail" + (bNewDirtyState ? " [*]" : "");
		}

		private void FormDefineHeaders_FormClosing(object sender, FormClosingEventArgs e)
		{
			//Form is closing

			//See if it needs saving
			if (_gbDirty)
			{
				//Show warning, only if the user initiated the closing
				if (e.CloseReason == CloseReason.UserClosing)
				{
					//Show a warning
					DialogResult resUsrChoice = MessageBox.Show("Your changes weren't saved yet.\n\n" +
						"Do you want to save changes before closing?",
						GlobalDefs.GlobDefs.gkstrAppNameFull, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);

					bool bAllowToClose = false;

					if (resUsrChoice == DialogResult.Yes)
					{
						//User chose to save changes
						if (saveData())
						{
							//Can continue
							bAllowToClose = true;
						}
					}
					else if (resUsrChoice == DialogResult.No)
					{
						//User chose not to save
						bAllowToClose = true;
					}


					if (!bAllowToClose)
					{
						//Prevent closing
						e.Cancel = true;
						return;
					}
				}
			}


			//Save size & location of this window
			bool bSaveProps = false;
			if (WindowState == FormWindowState.Normal)
			{
				//We're in normal state
				Properties.Settings.Default.OutlookHdrsWndLocation = Location;
				Properties.Settings.Default.OutlookHdrsWndSize = Size;
				Properties.Settings.Default.OutlookHdrsWndMax = false;

				bSaveProps = true;
			}
			else if (WindowState == FormWindowState.Maximized)
			{
				//We're maximized
				Properties.Settings.Default.OutlookHdrsWndLocation = RestoreBounds.Location;
				Properties.Settings.Default.OutlookHdrsWndSize = RestoreBounds.Size;
				Properties.Settings.Default.OutlookHdrsWndMax = true;

				bSaveProps = true;
			}

			if (true)
			{
				//Get sizes of the list headers
				StringCollection arrStrs = new StringCollection();
				foreach (ColumnHeader lvc in listHdrs.Columns)
				{
					arrStrs.Add(lvc.Width.ToString());
				}

				//Remember it
				Properties.Settings.Default.OutlookHdrsWndLVColSzs = arrStrs;
				bSaveProps = true;
			}


			if (bSaveProps)
			{
				try
				{
					//Make sure Outlook remembers these settings after a restart
					Properties.Settings.Default.Save();

				}
				catch(Exception ex)
				{
					//Failed -- can't show errors here as we're closing the app
					ThisAddIn.LogDiagnosticSpecError(1039, ex);
					SystemSounds.Exclamation.Play();
				}
			}
		}

		private void exitToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Close this app
			this.Close();
		}

		private void restoreSizeToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Restore window size and position
			//Also available through a short-cut Ctrl+Shift+R in case window position and size gets screwed up while saving or is moved too far off screen!

			Rectangle rcWrk = Screen.FromControl(this).WorkingArea;
			Size szWnd = this.MinimumSize;
			szWnd.Width = (int)(szWnd.Width * 1.3);
			szWnd.Height = (int)(szWnd.Height * 1.5);

			if (szWnd.Width < 500 || szWnd.Width > 1000)
			{
				szWnd.Width = 500;
			}

			if (szWnd.Height < 300 || szWnd.Height > 1000)
			{
				szWnd.Height = 300;
			}

			//Set position and size of our window
			Location = new Point(rcWrk.X + (rcWrk.Width - szWnd.Width) / 2, rcWrk.Y + (rcWrk.Height - szWnd.Height) / 2);
			Size = szWnd;
			WindowState = FormWindowState.Normal;

			//Restore column sizes in the list
			restoreHeaderListColumnSizes();
		}

		private void restoreHeaderListColumnSizes()
		{
			//Restore sizes of the header list columns
			int nW0, nW1, nW2;
			getDefaultHeaderListColumnSizes(out nW0, out nW1, out nW2);

			if (listHdrs.Columns.Count == 3)
			{
				listHdrs.Columns[0].Width = nW0;
				listHdrs.Columns[1].Width = nW1;
				listHdrs.Columns[2].Width = nW2;
			}
			else
			{
				//Error
				ThisAddIn.LogDiagnosticSpecError(1160);
				SystemSounds.Exclamation.Play();
			}
		}

		private void getDefaultHeaderListColumnSizes(out int nW0, out int nW1, out int nW2)
		{
			int nClientW = listHdrs.ClientSize.Width;
			int nWSb = System.Windows.Forms.SystemInformation.VerticalScrollBarWidth;
			int w = nClientW - nWSb;

			//Add two columns to the list
			nW0 = w * 10 / 100;

			//First column can't be too wide
			Size szM = TextRenderer.MeasureText("<ALL>++", this.Font);
			if (nW0 > szM.Width)
				nW0 = szM.Width;
			
			nW1 = w * 30 / 100;

			nW2 = w - nW0 - nW1;
		}

		private void restoreListColumnsToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Resore column sizes
			restoreHeaderListColumnSizes();
		}

		private void onlToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Show our help page
			showAppHelp();
		}

		static public void showAppHelp()
		{
			//See the type of installed used to install this add-in
			ThisAddIn.InstallerType installType = ThisAddIn.getCurrentInstallerType();
			int nInstallType = (int)installType;

			//Show help for this add-in
			ThisAddIn.openWebPage(ThisAddIn.kstrAppManualURL.
				Replace("#VER#", Uri.EscapeUriString(ThisAddIn.gAppVersion.getVersion())).
				Replace("#TYPE#", Uri.EscapeUriString(nInstallType.ToString())));
		}

		private void openEventLogToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Open diagnotic event log
			string strEventViewerPath = "";

			try
			{
				strEventViewerPath = Environment.SystemDirectory + "\\eventvwr.exe";

				System.Diagnostics.Process.Start(strEventViewerPath, "/c:Application");
			}
			catch (Exception ex)
			{
				//Failed
				ThisAddIn.LogDiagnosticSpecError(1042, "path: " + strEventViewerPath + " > " + ex);
				MessageBox.Show("ERROR: Failed to open Event Viewer: " + ex.ToString(), GlobalDefs.GlobDefs.gkstrAppNameFull, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void checkSelectedToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Check all items
			if (!checkUncheckAllHeaderItems(true, true))
			{
				//Failed
				ThisAddIn.LogDiagnosticSpecError(1128);
				SystemSounds.Exclamation.Play();
			}
		}

		private void uncheckSelectedToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Un-Check all items
			if (!checkUncheckAllHeaderItems(false, true))
			{
				//Failed
				ThisAddIn.LogDiagnosticSpecError(1129);
				SystemSounds.Exclamation.Play();
			}
		}

		private void comboBoxSendFmt_SelectedIndexChanged(object sender, EventArgs e)
		{
			//Selection in "Send emails in format:" changed

			if (_bAllowCatchListHdrsChanges)
			{
				EmlFmt_Item efi = (EmlFmt_Item)comboBoxSendFmt.SelectedItem;
				if (efi != null)
				{
					//Do we have our selected account
					if (_xmlNodeSelAcct != null)
					{
						//Set attribute
						XmlAttribute xAtt = _xmlNodeSelAcct.Attributes[kgStrNodeAttNm_SendAs];
						if (xAtt == null)
						{
							xAtt = _xmlHdrs.CreateAttribute(kgStrNodeAttNm_SendAs);
							xAtt.Value = ((int)efi.emlFmt).ToString();

							_xmlNodeSelAcct.Attributes.Append(xAtt);
						}
						else
						{
							xAtt.Value = ((int)efi.emlFmt).ToString();
						}

						//Update controls
						updateEmailFmtCtrls();

						//Set as "dirty"
						setDirty("9");
					}
					else
					{
						//Error
						ThisAddIn.LogDiagnosticSpecError(1131);
						SystemSounds.Exclamation.Play();
					}
				}
				else
				{
					//Error
					ThisAddIn.LogDiagnosticSpecError(1130);
					SystemSounds.Exclamation.Play();
				}
			}
		}

		private void updateEmailFmtCtrls()
		{
			//Update controls for CSS style
			bool bEnabled = false;
			EmlFmt_Item efi = (EmlFmt_Item)comboBoxSendFmt.SelectedItem;
			if (efi != null)
			{
				bEnabled = efi.emlFmt == EmailFmt.EML_FMT_SendAsHTML;
			}

			checkBoxCssStyle.Enabled = bEnabled;

			updateSetCssStyleCtrls();
		}

		private void updateSetCssStyleCtrls()
		{
			//Update controls for CSS style
			bool bEnabled = false;

			EmlFmt_Item efi = (EmlFmt_Item)comboBoxSendFmt.SelectedItem;
			if (efi != null)
			{
				bEnabled = efi.emlFmt == EmailFmt.EML_FMT_SendAsHTML;

				if (bEnabled)
				{
					if (checkBoxCssStyle.CheckState != CheckState.Checked)
						bEnabled = false;
				}
			}

			textBoxCssStyle.Enabled = bEnabled;
			buttonCtxMenuCssStyle.Enabled = bEnabled;
		}

		private void updateOverwriteMailClientNameCtrls()
		{
			//Update controls for overwrite mail client
			bool bEnable = checkBoxOverwriteMailClientNm.CheckState == CheckState.Checked;

			textBoxMailClientNm.Enabled = bEnable;
			buttonCtxMenuMailClientNm.Enabled = bEnable;

		}

		private void updateSuppressReceivedHaderCtrls()
		{
			//Update controls for suppress initial received header
			bool bEnable = checkBoxSuppressRcvd.CheckState == CheckState.Checked;

			textBoxSuppressRcvd.Enabled = bEnable;
			buttonCtxMenuSuppressRcvd.Enabled = bEnable;
		}

		private void updateTrackingPixelsCtrls()
		{
			//Update controls for tracking pixels
			bool bEnable = checkBoxRemTrPxl.CheckState == CheckState.Checked;

			textBoxTrckDomains.Enabled = bEnable;
			buttonCtxMenuTrckDomains.Enabled = bEnable;
		}


		public static string convertTriStateToDebugStr(TriState state)
		{
			switch (state)
			{
				case TriState.TS_True:
					return "Yes";
				case TriState.TS_False:
					return "No";
				case TriState.TS_Inherit:
					return "Inherit";
				default:
					return "<" + state.ToString() + ">";
			}
		}

		private static int convertThreeStateCheckToVal(CheckBox chkBox)
		{
			switch (chkBox.CheckState)
			{
				case CheckState.Checked:
					return 1;
				case CheckState.Unchecked:
					return 0;
				default:
					return 2;
			}
		}

		private static void setThreeStateCheckFromTriState(CheckBox chkBox, TriState val, bool bAllowTriState)
		{
			//Set tri-state checkbox
			switch (val)
			{
				case TriState.TS_True:
					chkBox.CheckState = CheckState.Checked;
					break;

				case TriState.TS_False:
					chkBox.CheckState = CheckState.Unchecked;
					break;

				default:
					chkBox.CheckState = bAllowTriState ? CheckState.Indeterminate : CheckState.Unchecked;
					break;
			}
		}


		private bool onTriStateCheckboxChanged(CheckBox chkBox, string strAttName, string strDebugDirty, int nOSErr1)
		{
			//RETURN:
			//		= true if changes were made - need to update controls
			bool bRes = false;

			//Checked status changed
			if (_bAllowCatchListHdrsChanges)
			{
				//Do we have our selected account
				if (_xmlNodeSelAcct != null)
				{
					//Value of the checkbox
					string strV = convertThreeStateCheckToVal(chkBox).ToString();

					//Set attribute
					XmlAttribute xAtt = _xmlNodeSelAcct.Attributes[strAttName];
					if (xAtt == null)
					{
						xAtt = _xmlHdrs.CreateAttribute(strAttName);
						xAtt.Value = strV;

						_xmlNodeSelAcct.Attributes.Append(xAtt);
					}
					else
					{
						xAtt.Value = strV;
					}

					//Update controls
					bRes = true;

					//Set as "dirty"
					setDirty(strDebugDirty);
				}
				else
				{
					//Error
					ThisAddIn.LogDiagnosticSpecError(nOSErr1);
					SystemSounds.Exclamation.Play();
				}
			}

			return bRes;
		}




		private void showContextMenuForButton(Button btnCtx, ContextMenuStrip ctxMenu)
		{
			//Show 'ctxMenu' context menu under the 'btnCtx' button

			//Get button position on the screen
			Point pntAt = btnCtx.Parent.PointToScreen(btnCtx.Location);
			pntAt.Y += btnCtx.Height;

			//Show context menu
			ctxMenu.Show(pntAt);
		}

		private void buttonCtxMenuMailClientNm_Click(object sender, EventArgs e)
		{
			//Show context menu
			showContextMenuForButton((Button)sender, contextMenuStripMailClientNm);
		}

		private void moreEmailClientsToolStripMenuItem_Click(object sender, EventArgs e)
		{
			ThisAddIn.openWebPage("https://user-agents.net/email-clients");
		}

		private void microsoftOutlookExpressToolStripMenuItem_Click(object sender, EventArgs e)
		{
			textBoxMailClientNm.Text = "Microsoft Outlook Express 6.0";
		}

		private void microsoftOutlook2007ToolStripMenuItem_Click(object sender, EventArgs e)
		{
			textBoxMailClientNm.Text = "Microsoft Office Outlook 12.0";
		}

		private void microsoftOutlook2010ToolStripMenuItem_Click(object sender, EventArgs e)
		{
			textBoxMailClientNm.Text = "Microsoft Outlook 14.0";
		}

		private void microsoftOutlook2013ToolStripMenuItem_Click(object sender, EventArgs e)
		{
			textBoxMailClientNm.Text = "Microsoft Outlook 15.0";
		}

		private void microsoftOutlook2016ToolStripMenuItem_Click(object sender, EventArgs e)
		{
			textBoxMailClientNm.Text = "Microsoft Outlook 16.0";
		}

		private void iPadMailToolStripMenuItem_Click(object sender, EventArgs e)
		{
			textBoxMailClientNm.Text = "iPad Mail (17F80)";
		}

		private void appleMailToolStripMenuItem_Click(object sender, EventArgs e)
		{
			textBoxMailClientNm.Text = "Apple Mail (13.2)";
		}

		private void thunderbirdToolStripMenuItem_Click(object sender, EventArgs e)
		{
			textBoxMailClientNm.Text = "Thunderbird 68.1";
		}

		private void buttonCtxMenuSuppressRcvd_Click(object sender, EventArgs e)
		{
			//Show context menu
			showContextMenuForButton((Button)sender, contextMenuStripSuppressRcvd);
		}

		private void amazonSESIPCurrentDateToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//from [192.168.1.124] (host-name.com [<ip>]) by pv50p00im-tydg10021701.me.com (Postfix) with ESMTPSA id 013098402CF for <dest-addr@gmail.com>; Sun, 12 Jul 2020 00:09:54 +0000 (UTC)
			textBoxSuppressRcvd.Text = "from [199.127.232.1] (b232-1.smtp-out.amazonses.com [199.127.232.1]) by auto-mailer (Postfix) for <#DEST_EMAIL_ADDR#>; #DATE_UTC#";
		}

		private void buttonCtxMenuCssStyle_Click(object sender, EventArgs e)
		{
			//Show context menu
			showContextMenuForButton((Button)sender, contextMenuStripCssStyle);

		}

		private void windows2000StyleToolStripMenuItem_Click(object sender, EventArgs e)
		{
			textBoxCssStyle.Text = "html,body,br,p,a {font: 15px Tahoma,\"Trebuchet MS\",Verdana,Arial,sans-serif;}";
		}

		private void windowsXPStyleToolStripMenuItem_Click(object sender, EventArgs e)
		{
			textBoxCssStyle.Text = "html,body,br,p,a {font: 15px \"Times New Roman\",Verdana,Tahoma,Arial,sans-serif;}";
		}

		private void windiwsVistaStyleToolStripMenuItem_Click(object sender, EventArgs e)
		{
			textBoxCssStyle.Text = "html,body,br,p,a {font: 15px Calibri,Candara,Consolas,Verdana,Arial,sans-serif;}";
		}

		private void windows10StyleToolStripMenuItem_Click(object sender, EventArgs e)
		{
			textBoxCssStyle.Text = "html,body,br,p,a {font: 14px \"Segoe UI\",\"Gill Sans Nova\",\"Georgia Pro\",Helvetica,Arial,Verdana,sans-serif;}";
		}

		private void universalStyle1ToolStripMenuItem_Click(object sender, EventArgs e)
		{
			textBoxCssStyle.Text = "html,body,br,p,a {font: 15px Arial,Helvetica,\"Times New Roman\",Times,Verdana,Tahoma;}";
		}

		private void universalStyle2ToolStripMenuItem_Click(object sender, EventArgs e)
		{
			textBoxCssStyle.Text = "html,body,br,p,a {font: 15px \"Times New Roman\",Times,Arial,Helvetica,Verdana,Tahoma;}";
		}

		private void universalStyle3ToolStripMenuItem_Click(object sender, EventArgs e)
		{
			textBoxCssStyle.Text = "html,body,br,p,a {font: 15px \"Courier New\",Courier,Helvetica,Arial,Verdana,sans-serif;}";
		}

		private void dennisbabkincomStyleToolStripMenuItem_Click(object sender, EventArgs e)
		{
			textBoxCssStyle.Text = "html,body,br,p,a {font: 14px \"Lucida Grande\",\"Lucida Sans Unicode\",Helvetica,Arial,Verdana,sans-serif;}";
		}

		private void buttonCtxMenuTrckDomains_Click(object sender, EventArgs e)
		{
			//Show context menu
			showContextMenuForButton((Button)sender, contextMenuStripTrckDomains);

		}

		private void clearToolStripMenuItem_Click(object sender, EventArgs e)
		{
			textBoxTrckDomains.Text = "";
		}

		private void amazonSESToolStripMenuItem_Click(object sender, EventArgs e)
		{
			addTrackingDomain("awstrack.me");

		}

		private void mailchimpToolStripMenuItem_Click(object sender, EventArgs e)
		{
			addTrackingDomain("list-manage.com");

		}

		private void sendinblueToolStripMenuItem_Click(object sender, EventArgs e)
		{
			addTrackingDomain("sendibt3.com");

		}


		private void addTrackingDomain(string strDomain)
		{
			//Add 'strDomain' to the list if it doesn't exist there
			string strCurrent = textBoxTrckDomains.Text;

			HashSet<string> arrDoms = ThisAddIn.getDomainListFromStr(strCurrent);

			//See if we have it already
			strDomain = strDomain.ToLower();

			if (!arrDoms.Contains(strDomain))
			{
				//Add it at the end
				strCurrent = strCurrent.Trim() + " " + strDomain;

				//And set it
				textBoxTrckDomains.Text = strCurrent.Trim();
			}
		}

		private void onTextBoxValChanged(TextBox txtBox, string strAttName, string strDebugDirty, int nOSErr1)
		{
			//Checked status changed
			if (_bAllowCatchListHdrsChanges)
			{
				//Do we have our selected account
				if (_xmlNodeSelAcct != null)
				{
					//Value of the checkbox
					string strV = txtBox.Text.Trim();

					//Set attribute
					XmlAttribute xAtt = _xmlNodeSelAcct.Attributes[strAttName];
					if (xAtt == null)
					{
						xAtt = _xmlHdrs.CreateAttribute(strAttName);
						xAtt.Value = strV;

						_xmlNodeSelAcct.Attributes.Append(xAtt);
					}
					else
					{
						xAtt.Value = strV;
					}

					//Set as "dirty"
					setDirty(strDebugDirty);
				}
				else
				{
					//Error
					ThisAddIn.LogDiagnosticSpecError(nOSErr1);
					SystemSounds.Exclamation.Play();
				}
			}

		}

		private void textBoxMailClientNm_TextChanged(object sender, EventArgs e)
		{
			//Text box value has changed
			onTextBoxValChanged((TextBox)sender, kgStrNodeAttNm_MailClientNm, "14", 1147);
		}

		private void textBoxSuppressRcvd_TextChanged(object sender, EventArgs e)
		{
			//Text box value has changed
			onTextBoxValChanged((TextBox)sender, kgStrNodeAttNm_RcvdHdr, "15", 1148);
		}

		private void textBoxCssStyle_TextChanged(object sender, EventArgs e)
		{
			//Text box value has changed
			onTextBoxValChanged((TextBox)sender, kgStrNodeAttNm_CSSStyles, "16", 1149);
		}

		private void textBoxTrckDomains_TextChanged(object sender, EventArgs e)
		{
			//Text box value has changed
			onTextBoxValChanged((TextBox)sender, kgStrNodeAttNm_TrackDoms, "17", 1150);
		}

		private void checkBoxOverwriteMailClientNm_CheckStateChanged(object sender, EventArgs e)
		{
			//Checked status changed
			if (onTriStateCheckboxChanged(checkBoxOverwriteMailClientNm, kgStrNodeAttNm_OverwriteMailClientNm, "13", 1146))
			{
				//Update controls
				updateOverwriteMailClientNameCtrls();
			}
		}

		private void checkBoxSuppressRcvd_CheckStateChanged(object sender, EventArgs e)
		{
			//Checked status changed
			if (onTriStateCheckboxChanged(checkBoxSuppressRcvd, kgStrNodeAttNm_SuppressRcvdHdr, "12", 1145))
			{
				//Update controls
				updateSuppressReceivedHaderCtrls();
			}
		}

		private void checkBoxCssStyle_CheckStateChanged(object sender, EventArgs e)
		{
			//Checked status changed
			if (onTriStateCheckboxChanged(checkBoxCssStyle, kgStrNodeAttNm_SetCSSStyles, "11", 1144))
			{
				//Update controls
				updateSetCssStyleCtrls();
			}
		}

		private void checkBoxRemTrPxl_CheckStateChanged(object sender, EventArgs e)
		{
			//Checked status changed
			if (onTriStateCheckboxChanged(checkBoxRemTrPxl, kgStrNodeAttNm_RemTrackPixel, "10", 1142))
			{
				//Update controls
				updateTrackingPixelsCtrls();
			}
		}

		private void checkForUpdatesToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Show our updates page

			//INFO: It is important to show to the end-user the same type of installer as was originally used by them!

			//See the type of installed used to install this add-in
			ThisAddIn.InstallerType installType = ThisAddIn.getCurrentInstallerType();
			int nInstallType = (int)installType;

			ThisAddIn.openWebPage(ThisAddIn.kstrAppUpdatesURL.
				Replace("#VER#", Uri.EscapeUriString(ThisAddIn.gAppVersion.getVersion())).
				Replace("#TYPE#", Uri.EscapeUriString(nInstallType.ToString())));
		}

		private void FormDefineHeaders_HelpRequested(object sender, HelpEventArgs hlpevent)
		{
			//Show help
			showAppHelp();
		}

		private void reportABugToolStripMenuItem_Click(object sender, EventArgs e)
		{
			//Send bug report

			//See the type of installed used to install this add-in
			ThisAddIn.InstallerType installType = ThisAddIn.getCurrentInstallerType();

			//Make report
			string strMsg = "MSI: " + installType.ToString() + ", Outlook v: " + Globals.ThisAddIn.Application.Version;

			ThisAddIn.openWebPage(ThisAddIn.kstrAppBugReportURL.
				Replace("#NAME#", Uri.EscapeUriString(GlobalDefs.GlobDefs.gkstrAppNameFull)).
				Replace("#VER#", Uri.EscapeUriString(ThisAddIn.gAppVersion.getVersion())).
				Replace("#MSG#", Uri.EscapeUriString(strMsg)));
		}



	}
}
