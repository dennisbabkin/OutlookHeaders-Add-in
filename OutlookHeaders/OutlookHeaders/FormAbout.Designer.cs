namespace OutlookHeaders
{
	partial class FormAbout
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.labelAppName = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.labelCpyrght = new System.Windows.Forms.Label();
			this.linkLabelDB = new System.Windows.Forms.LinkLabel();
			this.label2 = new System.Windows.Forms.Label();
			this.buttonOK = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.checkBoxLog = new System.Windows.Forms.CheckBox();
			this.textBoxOutput = new System.Windows.Forms.TextBox();
			this.pictureBox1 = new System.Windows.Forms.PictureBox();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
			this.SuspendLayout();
			// 
			// labelAppName
			// 
			this.labelAppName.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.labelAppName.Location = new System.Drawing.Point(93, 13);
			this.labelAppName.Name = "labelAppName";
			this.labelAppName.Size = new System.Drawing.Size(338, 18);
			this.labelAppName.TabIndex = 0;
			this.labelAppName.Text = "app name";
			this.labelAppName.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			// 
			// label1
			// 
			this.label1.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
			this.label1.Location = new System.Drawing.Point(12, 82);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(419, 24);
			this.label1.TabIndex = 1;
			this.label1.Text = "Outlook add-in for modifying mail headers && outbound emails.";
			this.label1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			// 
			// labelCpyrght
			// 
			this.labelCpyrght.Location = new System.Drawing.Point(93, 31);
			this.labelCpyrght.Name = "labelCpyrght";
			this.labelCpyrght.Size = new System.Drawing.Size(338, 18);
			this.labelCpyrght.TabIndex = 2;
			this.labelCpyrght.Text = "Copyright (C)";
			this.labelCpyrght.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			// 
			// linkLabelDB
			// 
			this.linkLabelDB.AutoSize = true;
			this.linkLabelDB.Location = new System.Drawing.Point(213, 49);
			this.linkLabelDB.Name = "linkLabelDB";
			this.linkLabelDB.Size = new System.Drawing.Size(93, 13);
			this.linkLabelDB.TabIndex = 3;
			this.linkLabelDB.TabStop = true;
			this.linkLabelDB.Text = "dennisbabkin.com";
			this.linkLabelDB.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelDB_LinkClicked);
			// 
			// label2
			// 
			this.label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.label2.Location = new System.Drawing.Point(15, 74);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(419, 2);
			this.label2.TabIndex = 4;
			// 
			// buttonOK
			// 
			this.buttonOK.Location = new System.Drawing.Point(355, 291);
			this.buttonOK.Name = "buttonOK";
			this.buttonOK.Size = new System.Drawing.Size(75, 23);
			this.buttonOK.TabIndex = 5;
			this.buttonOK.Text = "OK";
			this.buttonOK.UseVisualStyleBackColor = true;
			this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.checkBoxLog);
			this.groupBox1.Controls.Add(this.textBoxOutput);
			this.groupBox1.Location = new System.Drawing.Point(15, 109);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(416, 175);
			this.groupBox1.TabIndex = 7;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Debugging Output";
			// 
			// checkBoxLog
			// 
			this.checkBoxLog.AutoSize = true;
			this.checkBoxLog.Location = new System.Drawing.Point(18, 147);
			this.checkBoxLog.Name = "checkBoxLog";
			this.checkBoxLog.Size = new System.Drawing.Size(208, 17);
			this.checkBoxLog.TabIndex = 8;
			this.checkBoxLog.Text = "Log diagnotic events when mail is sent";
			this.checkBoxLog.UseVisualStyleBackColor = true;
			// 
			// textBoxOutput
			// 
			this.textBoxOutput.ForeColor = System.Drawing.SystemColors.WindowFrame;
			this.textBoxOutput.Location = new System.Drawing.Point(18, 23);
			this.textBoxOutput.Multiline = true;
			this.textBoxOutput.Name = "textBoxOutput";
			this.textBoxOutput.ReadOnly = true;
			this.textBoxOutput.Size = new System.Drawing.Size(381, 118);
			this.textBoxOutput.TabIndex = 7;
			// 
			// pictureBox1
			// 
			this.pictureBox1.Image = global::OutlookHeaders.Properties.Resources.olh_main_icn_64x64;
			this.pictureBox1.Location = new System.Drawing.Point(27, 13);
			this.pictureBox1.Name = "pictureBox1";
			this.pictureBox1.Size = new System.Drawing.Size(54, 49);
			this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.pictureBox1.TabIndex = 8;
			this.pictureBox1.TabStop = false;
			// 
			// FormAbout
			// 
			this.AcceptButton = this.buttonOK;
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(443, 325);
			this.Controls.Add(this.pictureBox1);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.buttonOK);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.linkLabelDB);
			this.Controls.Add(this.labelCpyrght);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.labelAppName);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "FormAbout";
			this.ShowIcon = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "About This Add-in";
			this.Load += new System.EventHandler(this.FormAbout_Load);
			this.groupBox1.ResumeLayout(false);
			this.groupBox1.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label labelAppName;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label labelCpyrght;
		private System.Windows.Forms.LinkLabel linkLabelDB;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button buttonOK;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.TextBox textBoxOutput;
		private System.Windows.Forms.CheckBox checkBoxLog;
		private System.Windows.Forms.PictureBox pictureBox1;
	}
}