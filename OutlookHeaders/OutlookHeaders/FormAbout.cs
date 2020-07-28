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

using System.Configuration;




namespace OutlookHeaders
{
	public partial class FormAbout : Form
	{
		public FormAbout()
		{
			InitializeComponent();
		}

		private void FormAbout_Load(object sender, EventArgs e)
		{
			//First loading

			//Get app installer type
			ThisAddIn.InstallerType instType = ThisAddIn.getCurrentInstallerType();
			string strInst;
			switch (instType)
			{
				case ThisAddIn.InstallerType.INST_PerUser:
					strInst = "Per-User";
					break;

				case ThisAddIn.InstallerType.INST_AllUsers:
					strInst = "All-Users";
					break;

				case ThisAddIn.InstallerType.INST_Bundled:
					strInst = "Bundled";
					break;

				case ThisAddIn.InstallerType.INST_Unknown:
					strInst = "?";
					break;

				default:
					ThisAddIn.LogDiagnosticSpecError(1158, "v=" + instType.ToString());
					strInst = instType.ToString();
					break;
			}

			labelAppName.Text = GlobalDefs.GlobDefs.gkstrAppNameFull + " v." + ThisAddIn.gAppVersion.getVersion() + " (" + strInst + ")";

			//Copyright
			int nYear = DateTime.Now.Year;
			labelCpyrght.Text = "Copyright (C) 2020" + (nYear > 2020 ? "-" + nYear.ToString() : "");


			//Get path for settings
			string strSettingsPath = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal).FilePath;


			textBoxOutput.Text = "ConfigFilePath=\"" + strSettingsPath + "\"\r\n" + 
				"ExecFldr=\"" + AppDomain.CurrentDomain.BaseDirectory + "\"";

			textBoxOutput.ForeColor = SystemColors.GrayText;
			textBoxOutput.BackColor = SystemColors.ButtonFace;		//I'm not sure why we need to do this?

			//Set checkbox for logging sent emails
			checkBoxLog.Checked = Properties.Settings.Default.OutlookHdrsLogWhenSending;
		}

		private void buttonOK_Click(object sender, EventArgs e)
		{
			//Save any changes
			try
			{
				Properties.Settings.Default.OutlookHdrsLogWhenSending = checkBoxLog.Checked;

				//And save them
				Properties.Settings.Default.Save();
			}
			catch(Exception ex)
			{
				//Failed
				ThisAddIn.LogDiagnosticSpecError(1041, ex);
				MessageBox.Show("ERROR: Failed to save changes: " + ex.ToString(), GlobalDefs.GlobDefs.gkstrAppNameFull, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}

			//Close this window
			this.Close();
		}

		private void linkLabelDB_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
		{
			//Redirect to the website of the developer
			ThisAddIn.openWebPage("https://dennisbabkin.com/");

		}
	}
}
