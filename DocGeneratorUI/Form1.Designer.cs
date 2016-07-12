namespace DocGeneratorUI
	{
	partial class Form1
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
			if(disposing && (components != null))
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
			this.pictureBox1 = new System.Windows.Forms.PictureBox();
			this.button1 = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.groupBoxSpecific = new System.Windows.Forms.GroupBox();
			this.maskedTextBox1 = new System.Windows.Forms.MaskedTextBox();
			this.button2 = new System.Windows.Forms.Button();
			this.label2 = new System.Windows.Forms.Label();
			this.statusStrip1 = new System.Windows.Forms.StatusStrip();
			this.toolStripStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
			this.toolStripProgressBar = new System.Windows.Forms.ToolStripProgressBar();
			this.buttonSendEmail = new System.Windows.Forms.Button();
			this.textBoxEmail = new System.Windows.Forms.TextBox();
			this.labelEmail = new System.Windows.Forms.Label();
			this.groupBoxEmails = new System.Windows.Forms.GroupBox();
			this.radioError = new System.Windows.Forms.RadioButton();
			this.radioWarning = new System.Windows.Forms.RadioButton();
			this.radioInformation = new System.Windows.Forms.RadioButton();
			this.buttonSendTechnicalEmail = new System.Windows.Forms.Button();
			this.buttonSendUserErrorEmail = new System.Windows.Forms.Button();
			this.comboBoxPlatform = new System.Windows.Forms.ComboBox();
			this.labelPlatform = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
			this.groupBoxSpecific.SuspendLayout();
			this.statusStrip1.SuspendLayout();
			this.groupBoxEmails.SuspendLayout();
			this.SuspendLayout();
			// 
			// pictureBox1
			// 
			this.pictureBox1.Cursor = System.Windows.Forms.Cursors.No;
			this.pictureBox1.Image = global::DocGeneratorUI.Properties.Resources.DocGenerator;
			this.pictureBox1.Location = new System.Drawing.Point(12, 12);
			this.pictureBox1.Name = "pictureBox1";
			this.pictureBox1.Size = new System.Drawing.Size(81, 81);
			this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
			this.pictureBox1.TabIndex = 1;
			this.pictureBox1.TabStop = false;
			// 
			// button1
			// 
			this.button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
			this.button1.Cursor = System.Windows.Forms.Cursors.Hand;
			this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.button1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(105)))), ((int)(((byte)(190)))), ((int)(((byte)(42)))));
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Location = new System.Drawing.Point(22, 140);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(381, 39);
			this.button1.TabIndex = 0;
			this.button1.Text = "Generate All Unprocessed Document Collections";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.UseVisualStyleBackColor = true;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// label1
			// 
			this.label1.AllowDrop = true;
			this.label1.Cursor = System.Windows.Forms.Cursors.No;
			this.label1.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
			this.label1.Location = new System.Drawing.Point(100, 12);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(416, 42);
			this.label1.TabIndex = 2;
			this.label1.Text = "This is the User Inteface to the Document Generator, which makes it possible to e" +
    "xecute the DocGenenrator Core DLL on a local computer.";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label1.UseCompatibleTextRendering = true;
			// 
			// groupBoxSpecific
			// 
			this.groupBoxSpecific.Controls.Add(this.maskedTextBox1);
			this.groupBoxSpecific.Controls.Add(this.button2);
			this.groupBoxSpecific.Controls.Add(this.label2);
			this.groupBoxSpecific.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.groupBoxSpecific.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.groupBoxSpecific.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
			this.groupBoxSpecific.Location = new System.Drawing.Point(12, 185);
			this.groupBoxSpecific.Name = "groupBoxSpecific";
			this.groupBoxSpecific.Size = new System.Drawing.Size(407, 92);
			this.groupBoxSpecific.TabIndex = 2;
			this.groupBoxSpecific.TabStop = false;
			this.groupBoxSpecific.Text = "Generate specific Document Collection";
			// 
			// maskedTextBox1
			// 
			this.maskedTextBox1.AsciiOnly = true;
			this.maskedTextBox1.BeepOnError = true;
			this.maskedTextBox1.Location = new System.Drawing.Point(152, 21);
			this.maskedTextBox1.Mask = "0000000";
			this.maskedTextBox1.Name = "maskedTextBox1";
			this.maskedTextBox1.Size = new System.Drawing.Size(58, 20);
			this.maskedTextBox1.TabIndex = 5;
			// 
			// button2
			// 
			this.button2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
			this.button2.Cursor = System.Windows.Forms.Cursors.Hand;
			this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.button2.ForeColor = System.Drawing.Color.RoyalBlue;
			this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
			this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button2.Location = new System.Drawing.Point(10, 47);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(381, 39);
			this.button2.TabIndex = 4;
			this.button2.Text = "Generate the entered Document Collection";
			this.button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button2.UseVisualStyleBackColor = true;
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label2.Location = new System.Drawing.Point(7, 24);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(139, 15);
			this.label2.TabIndex = 0;
			this.label2.Text = "Document Collection ID:";
			// 
			// statusStrip1
			// 
			this.statusStrip1.AutoSize = false;
			this.statusStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Visible;
			this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel,
            this.toolStripProgressBar});
			this.statusStrip1.Location = new System.Drawing.Point(0, 541);
			this.statusStrip1.Name = "statusStrip1";
			this.statusStrip1.Size = new System.Drawing.Size(524, 22);
			this.statusStrip1.TabIndex = 6;
			this.statusStrip1.Text = "statusStrip1";
			// 
			// toolStripStatusLabel
			// 
			this.toolStripStatusLabel.AutoSize = false;
			this.toolStripStatusLabel.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
			this.toolStripStatusLabel.ForeColor = System.Drawing.Color.Green;
			this.toolStripStatusLabel.Margin = new System.Windows.Forms.Padding(1, 2, 1, 2);
			this.toolStripStatusLabel.Name = "toolStripStatusLabel";
			this.toolStripStatusLabel.Size = new System.Drawing.Size(250, 18);
			this.toolStripStatusLabel.Text = "Ready....";
			this.toolStripStatusLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// toolStripProgressBar
			// 
			this.toolStripProgressBar.Enabled = false;
			this.toolStripProgressBar.Margin = new System.Windows.Forms.Padding(1, 2, 1, 2);
			this.toolStripProgressBar.Maximum = 6;
			this.toolStripProgressBar.Name = "toolStripProgressBar";
			this.toolStripProgressBar.Size = new System.Drawing.Size(225, 18);
			this.toolStripProgressBar.Step = 1;
			this.toolStripProgressBar.Visible = false;
			// 
			// buttonSendEmail
			// 
			this.buttonSendEmail.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
			this.buttonSendEmail.Cursor = System.Windows.Forms.Cursors.Hand;
			this.buttonSendEmail.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.buttonSendEmail.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.buttonSendEmail.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(105)))), ((int)(((byte)(190)))), ((int)(((byte)(42)))));
			this.buttonSendEmail.Image = ((System.Drawing.Image)(resources.GetObject("buttonSendEmail.Image")));
			this.buttonSendEmail.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.buttonSendEmail.Location = new System.Drawing.Point(110, 64);
			this.buttonSendEmail.Name = "buttonSendEmail";
			this.buttonSendEmail.Size = new System.Drawing.Size(382, 51);
			this.buttonSendEmail.TabIndex = 4;
			this.buttonSendEmail.Text = "Send an example of a Successful User E-mail";
			this.buttonSendEmail.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.buttonSendEmail.UseVisualStyleBackColor = true;
			this.buttonSendEmail.Click += new System.EventHandler(this.buttonSendEmail_Click);
			// 
			// textBoxEmail
			// 
			this.textBoxEmail.Location = new System.Drawing.Point(110, 39);
			this.textBoxEmail.MaxLength = 100;
			this.textBoxEmail.Name = "textBoxEmail";
			this.textBoxEmail.Size = new System.Drawing.Size(382, 20);
			this.textBoxEmail.TabIndex = 5;
			// 
			// labelEmail
			// 
			this.labelEmail.CausesValidation = false;
			this.labelEmail.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.labelEmail.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.labelEmail.Location = new System.Drawing.Point(7, 16);
			this.labelEmail.Name = "labelEmail";
			this.labelEmail.Size = new System.Drawing.Size(491, 20);
			this.labelEmail.TabIndex = 6;
			this.labelEmail.Text = "Enter the e-mail address of the recipeit of the sample message must be send:";
			this.labelEmail.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBoxEmails
			// 
			this.groupBoxEmails.Controls.Add(this.radioError);
			this.groupBoxEmails.Controls.Add(this.radioWarning);
			this.groupBoxEmails.Controls.Add(this.radioInformation);
			this.groupBoxEmails.Controls.Add(this.buttonSendTechnicalEmail);
			this.groupBoxEmails.Controls.Add(this.buttonSendUserErrorEmail);
			this.groupBoxEmails.Controls.Add(this.textBoxEmail);
			this.groupBoxEmails.Controls.Add(this.buttonSendEmail);
			this.groupBoxEmails.Controls.Add(this.labelEmail);
			this.groupBoxEmails.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.groupBoxEmails.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.groupBoxEmails.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
			this.groupBoxEmails.Location = new System.Drawing.Point(12, 285);
			this.groupBoxEmails.Name = "groupBoxEmails";
			this.groupBoxEmails.Size = new System.Drawing.Size(504, 245);
			this.groupBoxEmails.TabIndex = 7;
			this.groupBoxEmails.TabStop = false;
			this.groupBoxEmails.Text = "Test HTML E-mails";
			// 
			// radioError
			// 
			this.radioError.AutoSize = true;
			this.radioError.Location = new System.Drawing.Point(7, 217);
			this.radioError.Name = "radioError";
			this.radioError.Size = new System.Drawing.Size(52, 17);
			this.radioError.TabIndex = 11;
			this.radioError.Text = "Error";
			this.radioError.UseVisualStyleBackColor = true;
			// 
			// radioWarning
			// 
			this.radioWarning.AutoSize = true;
			this.radioWarning.Location = new System.Drawing.Point(7, 199);
			this.radioWarning.Name = "radioWarning";
			this.radioWarning.Size = new System.Drawing.Size(72, 17);
			this.radioWarning.TabIndex = 10;
			this.radioWarning.Text = "Warning";
			this.radioWarning.UseVisualStyleBackColor = true;
			// 
			// radioInformation
			// 
			this.radioInformation.AutoSize = true;
			this.radioInformation.Checked = true;
			this.radioInformation.Location = new System.Drawing.Point(7, 181);
			this.radioInformation.Name = "radioInformation";
			this.radioInformation.Size = new System.Drawing.Size(88, 17);
			this.radioInformation.TabIndex = 9;
			this.radioInformation.TabStop = true;
			this.radioInformation.Text = "Information";
			this.radioInformation.UseVisualStyleBackColor = true;
			// 
			// buttonSendTechnicalEmail
			// 
			this.buttonSendTechnicalEmail.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
			this.buttonSendTechnicalEmail.Cursor = System.Windows.Forms.Cursors.Hand;
			this.buttonSendTechnicalEmail.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.buttonSendTechnicalEmail.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.buttonSendTechnicalEmail.ForeColor = System.Drawing.Color.Red;
			this.buttonSendTechnicalEmail.Image = ((System.Drawing.Image)(resources.GetObject("buttonSendTechnicalEmail.Image")));
			this.buttonSendTechnicalEmail.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.buttonSendTechnicalEmail.Location = new System.Drawing.Point(111, 181);
			this.buttonSendTechnicalEmail.Name = "buttonSendTechnicalEmail";
			this.buttonSendTechnicalEmail.Size = new System.Drawing.Size(382, 54);
			this.buttonSendTechnicalEmail.TabIndex = 8;
			this.buttonSendTechnicalEmail.Text = "Send an example of a Technical Support E-mail";
			this.buttonSendTechnicalEmail.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.buttonSendTechnicalEmail.UseVisualStyleBackColor = true;
			this.buttonSendTechnicalEmail.Click += new System.EventHandler(this.buttonSendTechnicalEmail_Click);
			// 
			// buttonSendUserErrorEmail
			// 
			this.buttonSendUserErrorEmail.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
			this.buttonSendUserErrorEmail.Cursor = System.Windows.Forms.Cursors.Hand;
			this.buttonSendUserErrorEmail.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.buttonSendUserErrorEmail.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.buttonSendUserErrorEmail.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
			this.buttonSendUserErrorEmail.Image = ((System.Drawing.Image)(resources.GetObject("buttonSendUserErrorEmail.Image")));
			this.buttonSendUserErrorEmail.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.buttonSendUserErrorEmail.Location = new System.Drawing.Point(110, 121);
			this.buttonSendUserErrorEmail.Name = "buttonSendUserErrorEmail";
			this.buttonSendUserErrorEmail.Size = new System.Drawing.Size(382, 54);
			this.buttonSendUserErrorEmail.TabIndex = 7;
			this.buttonSendUserErrorEmail.Text = "Send an example of a User error E-mail";
			this.buttonSendUserErrorEmail.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.buttonSendUserErrorEmail.UseVisualStyleBackColor = true;
			this.buttonSendUserErrorEmail.Click += new System.EventHandler(this.buttonSendUserErrorEmail_Click);
			// 
			// comboBoxPlatform
			// 
			this.comboBoxPlatform.AutoCompleteCustomSource.AddRange(new string[] {
            "Development Environment",
            "Quality Assurance Environment",
            "Production Environment"});
			this.comboBoxPlatform.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
			this.comboBoxPlatform.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
			this.comboBoxPlatform.Cursor = System.Windows.Forms.Cursors.Hand;
			this.comboBoxPlatform.Items.AddRange(new object[] {
            "Development Environment",
            "Quality Assurance Environment",
            "Production Environment"});
			this.comboBoxPlatform.Location = new System.Drawing.Point(110, 107);
			this.comboBoxPlatform.MaxDropDownItems = 3;
			this.comboBoxPlatform.Name = "comboBoxPlatform";
			this.comboBoxPlatform.Size = new System.Drawing.Size(293, 21);
			this.comboBoxPlatform.TabIndex = 0;
			this.comboBoxPlatform.Text = "Select the platform...";
			this.comboBoxPlatform.SelectedIndexChanged += new System.EventHandler(this.comboBoxPlatform_SelectedIndexChanged);
			// 
			// labelPlatform
			// 
			this.labelPlatform.CausesValidation = false;
			this.labelPlatform.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.labelPlatform.Location = new System.Drawing.Point(31, 110);
			this.labelPlatform.Name = "labelPlatform";
			this.labelPlatform.Size = new System.Drawing.Size(73, 20);
			this.labelPlatform.TabIndex = 9;
			this.labelPlatform.Text = "Platform:";
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(524, 563);
			this.Controls.Add(this.comboBoxPlatform);
			this.Controls.Add(this.labelPlatform);
			this.Controls.Add(this.groupBoxEmails);
			this.Controls.Add(this.statusStrip1);
			this.Controls.Add(this.groupBoxSpecific);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.pictureBox1);
			this.Controls.Add(this.button1);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "Form1";
			this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
			this.Text = "DocGenerator User Interface";
			this.Load += new System.EventHandler(this.Form1_Load);
			((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
			this.groupBoxSpecific.ResumeLayout(false);
			this.groupBoxSpecific.PerformLayout();
			this.statusStrip1.ResumeLayout(false);
			this.statusStrip1.PerformLayout();
			this.groupBoxEmails.ResumeLayout(false);
			this.groupBoxEmails.PerformLayout();
			this.ResumeLayout(false);

			}

		#endregion

		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.PictureBox pictureBox1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox groupBoxSpecific;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.MaskedTextBox maskedTextBox1;
		private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar;
		public System.Windows.Forms.StatusStrip statusStrip1;
		public System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel;
		private System.Windows.Forms.Button buttonSendEmail;
		private System.Windows.Forms.TextBox textBoxEmail;
		private System.Windows.Forms.Label labelEmail;
		private System.Windows.Forms.GroupBox groupBoxEmails;
		private System.Windows.Forms.Button buttonSendTechnicalEmail;
		private System.Windows.Forms.Button buttonSendUserErrorEmail;
		private System.Windows.Forms.RadioButton radioError;
		private System.Windows.Forms.RadioButton radioWarning;
		private System.Windows.Forms.RadioButton radioInformation;
		private System.Windows.Forms.ComboBox comboBoxPlatform;
		private System.Windows.Forms.Label labelPlatform;
		}
	}

