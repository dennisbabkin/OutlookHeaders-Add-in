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
using System.Xml;

using System.Media;



namespace OutlookHeaders
{
    public partial class FormEditHeader : Form
    {
        public class HeaderType
        {
			public XmlDocument xmlDoc = null;		//XML document used for storing the data
			public XmlNode xmlParNode = null;		//XML parent node where adding or editing item in
			public XmlNode xmlNode = null;			//XML node in 'xmlNode' of the item being edited, or null if adding new. On the output, this is the node that was added.
        }

        public HeaderType hdrType = new HeaderType();       //[in]/[out] receives and sets the values selected


        public FormEditHeader()
        {
            InitializeComponent();
        }

        private void FormEditHeader_Load(object sender, EventArgs e)
        {
            //Form loading
			bool bAddNew = hdrType.xmlNode == null;

			//Set min size
			this.MinimumSize = this.Size;

            if (bAddNew)
            {
                Text = "Add New Mail Header";
                buttonOK.Text = "&Add";
            }
            else
            {
                Text = "Edit Mail Header";
                buttonOK.Text = "&Save";

				//Get values from XML node
				FormDefineHeaders.ItemInfo ii = FormDefineHeaders.getHeaderInfoFromNode(hdrType.xmlNode);
				if (ii != null)
				{
					textBoxName.Text = ii.strName.Trim();
					textBoxValue.Text = ii.strVal.Trim();
				}
				else
				{
					//Error
					ThisAddIn.LogDiagnosticSpecError(1040);
					SystemSounds.Exclamation.Play();
				}
            }
        }

		public static Dictionary<char, string> validateFieldName(string strName, int nMaxNumErrors)
		{
			//Check 'strName' for invalid characters
			//'nMaxNumErrors' = maximum number of errors to collect, or 0 to collect all
			//RETURN:
			//		= Dictionary of invalid chars found
			Dictionary<char, string> dicRes = new Dictionary<char, string>();

			//SOURCE:
			//		https://tools.ietf.org/html/rfc5322#section-2.2
			//		https://tools.ietf.org/html/rfc2047
			//		
			if (!string.IsNullOrEmpty(strName))
			{
				//Bad chars in the name
				Dictionary<char, string> kdicBadChars = new Dictionary<char, string>()
				{
					{'\t', "Tab"},
					{'\n', "Line feed"},
					{'\r', "Carriage return"},
					{' ', "Space"},
				};

				string strCh;

				int nLn = strName.Length;
				for (int i = 0; i < nLn; i++)
				{
					char c = strName[i];

					if (c >= 33 && c <= 126 && c != ':')
					{
						//This char is good
					}
					else
					{
						//Bad char, pick the name for it
						if (!kdicBadChars.TryGetValue(c, out strCh))
						{
							if (c < 32)
								strCh = "Character with code " + ((int)c).ToString();
							else
								strCh = "Character '" + c + "' with code " + ((int)c).ToString();
						}

						dicRes[c] = strCh;

						if (nMaxNumErrors > 0)
						{
							if (dicRes.Count >= nMaxNumErrors)
								break;
						}
					}
				}
			}

			return dicRes;
		}

		public static Dictionary<char, string> validateFieldValue(string strVal, int nMaxNumErrors)
		{
			//Check 'strVal' for invalid characters
			//'nMaxNumErrors' = maximum number of errors to collect, or 0 to collect all
			//RETURN:
			//		= Dictionary of invalid chars found
			Dictionary<char, string> dicRes = new Dictionary<char, string>();

			//SOURCE:
			//		https://tools.ietf.org/html/rfc5322#section-2.2
			//		https://tools.ietf.org/html/rfc2047
			//		
			if (!string.IsNullOrEmpty(strVal))
			{
				//Bad chars in values (or fields)
				Dictionary<char, string> kdicBadChars = new Dictionary<char, string>()
				{
					{'\n', "Line feed"},
					{'\r', "Carriage return"},
				};

				string strCh;

				int nLn = strVal.Length;
				for (int i = 0; i < nLn; i++)
				{
					char c = strVal[i];

					if ((c >= 32 && c <= 126) || c == '\t')
					{
						//This char is good
					}
					else
					{
						//Bad char, pick the name for it
						if (!kdicBadChars.TryGetValue(c, out strCh))
						{
							if (c < 32)
								strCh = "Character with code " + ((int)c).ToString();
							else
								strCh = "Character '" + c + "' with code " + ((int)c).ToString();
						}

						dicRes[c] = strCh;

						if (nMaxNumErrors > 0)
						{
							if (dicRes.Count >= nMaxNumErrors)
								break;
						}
					}
				}
			}

			return dicRes;
		}

		private static string getListOfBadChars(Dictionary<char, string> dicBadChars)
		{
			//RETURN: = String with bad chars on each line
			string str = "";

			foreach (KeyValuePair<char, string> kvp in dicBadChars)
			{
				str += " - " + kvp.Value + "\n";
			}

			return str;
		}


        private void buttonOK_Click(object sender, EventArgs e)
        {
            //OK button clicked
			bool bAddNew = hdrType.xmlNode == null;

            string strName = textBoxName.Text.Trim();
            string strValue = textBoxValue.Text.Trim();

			//See if user specified the http header name
            if (string.IsNullOrEmpty(strName))
            {
				MessageBox.Show("Header name is required.", GlobalDefs.GlobDefs.gkstrAppNameFull, MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBoxName.Focus();
                return;
            }

			//Check for illegal chars in the name
			Dictionary<char, string> dicValName = validateFieldName(strName, 5);
			if (dicValName.Count > 0)
			{
				//Issue a warning
				if (MessageBox.Show("Header field name contains the following illegal characters:\n\n" +
					getListOfBadChars(dicValName) + "\n" +
					"Do you want to use that name?\n\n" +
					"(If you continue, your mail server may reject this header.)"
					,
					GlobalDefs.GlobDefs.gkstrAppNameFull, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes)
				{
					textBoxName.Focus();
					return;
				}
			}

			//Check for illegal chars in the value
			dicValName = validateFieldValue(strValue, 5);
			if (dicValName.Count > 0)
			{
				//Issue a warning
				if (MessageBox.Show("Header field body contains the following illegal characters:\n\n" +
					getListOfBadChars(dicValName) + "\n" +
					"Do you want to use it?\n\n" +
					"(If you continue, your mail server may reject this header.)"
					,
					GlobalDefs.GlobDefs.gkstrAppNameFull, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes)
				{
					textBoxValue.Focus();
					return;
				}
			}


			if (bAddNew)
			{
				//Add this pair to our XML node
				XmlNode xNew = hdrType.xmlDoc.CreateElement(FormDefineHeaders.kgStrNodeNm_Header);

				//Name
				XmlAttribute xAttNm = hdrType.xmlDoc.CreateAttribute(FormDefineHeaders.kgStrNodeAttNm_Name);
				xAttNm.Value = strName;
				xNew.Attributes.Append(xAttNm);

				//Enabled by default
				XmlAttribute xAttEnabled = hdrType.xmlDoc.CreateAttribute(FormDefineHeaders.kgStrNodeAttNm_Enabled);
				xAttEnabled.Value = "1";
				xNew.Attributes.Append(xAttEnabled);

				//Value
				xNew.InnerText = strValue;

				hdrType.xmlParNode.AppendChild(xNew);

				//Return this node back
				hdrType.xmlNode = xNew;
			}
			else
			{
				//Adjust existing header
				XmlAttribute xAttNm = hdrType.xmlNode.Attributes[FormDefineHeaders.kgStrNodeAttNm_Name];
				if (xAttNm != null)
				{
					xAttNm.Value = strName;
				}
				else
				{
					//Add new
					xAttNm = hdrType.xmlDoc.CreateAttribute(FormDefineHeaders.kgStrNodeAttNm_Name);
					xAttNm.Value = strName;

					hdrType.xmlNode.Attributes.Append(xAttNm);
					xAttNm = hdrType.xmlNode.Attributes[FormDefineHeaders.kgStrNodeAttNm_Name];
				}

				//Value
				hdrType.xmlNode.InnerText = strValue;
			}


            //Close with OK result
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

    }
}
