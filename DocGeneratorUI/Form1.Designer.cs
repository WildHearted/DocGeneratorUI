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
			this.components = new System.ComponentModel.Container();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
			this.pictureBox1 = new System.Windows.Forms.PictureBox();
			this.buttonGenerateAll = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.groupBoxSpecific = new System.Windows.Forms.GroupBox();
			this.textBoxDocumentCollection = new System.Windows.Forms.TextBox();
			this.buttonGenerateDocCollection = new System.Windows.Forms.Button();
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
			this.buttonTestHTML = new System.Windows.Forms.Button();
			this.buttonGenerateTestDoc = new System.Windows.Forms.Button();
			this.buttonTestSomething = new System.Windows.Forms.Button();
			this.groupBoxPlatform = new System.Windows.Forms.GroupBox();
			this.radioButtonDEV = new System.Windows.Forms.RadioButton();
			this.radioButtonQA = new System.Windows.Forms.RadioButton();
			this.radioButtonPROD = new System.Windows.Forms.RadioButton();
			this.toolTip = new System.Windows.Forms.ToolTip(this.components);
			((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
			this.groupBoxSpecific.SuspendLayout();
			this.statusStrip1.SuspendLayout();
			this.groupBoxEmails.SuspendLayout();
			this.groupBoxPlatform.SuspendLayout();
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
			// buttonGenerateAll
			// 
			this.buttonGenerateAll.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
			this.buttonGenerateAll.Cursor = System.Windows.Forms.Cursors.Hand;
			this.buttonGenerateAll.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.buttonGenerateAll.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.buttonGenerateAll.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(105)))), ((int)(((byte)(190)))), ((int)(((byte)(42)))));
			this.buttonGenerateAll.Image = ((System.Drawing.Image)(resources.GetObject("buttonGenerateAll.Image")));
			this.buttonGenerateAll.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.buttonGenerateAll.Location = new System.Drawing.Point(12, 176);
			this.buttonGenerateAll.Name = "buttonGenerateAll";
			this.buttonGenerateAll.Size = new System.Drawing.Size(381, 39);
			this.buttonGenerateAll.TabIndex = 0;
			this.buttonGenerateAll.Text = "Generate All Unprocessed Document Collections";
			this.buttonGenerateAll.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.buttonGenerateAll.UseVisualStyleBackColor = true;
			this.buttonGenerateAll.Click += new System.EventHandler(this.buttonGenerateAll_Click);
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
			this.groupBoxSpecific.Controls.Add(this.textBoxDocumentCollection);
			this.groupBoxSpecific.Controls.Add(this.buttonGenerateDocCollection);
			this.groupBoxSpecific.Controls.Add(this.label2);
			this.groupBoxSpecific.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.groupBoxSpecific.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.groupBoxSpecific.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
			this.groupBoxSpecific.Location = new System.Drawing.Point(12, 238);
			this.groupBoxSpecific.Name = "groupBoxSpecific";
			this.groupBoxSpecific.Size = new System.Drawing.Size(407, 92);
			this.groupBoxSpecific.TabIndex = 2;
			this.groupBoxSpecific.TabStop = false;
			this.groupBoxSpecific.Text = "Generate specific Document Collection";
			// 
			// textBoxDocumentCollection
			// 
			this.textBoxDocumentCollection.Location = new System.Drawing.Point(152, 22);
			this.textBoxDocumentCollection.MaxLength = 10;
			this.textBoxDocumentCollection.Name = "textBoxDocumentCollection";
			this.textBoxDocumentCollection.Size = new System.Drawing.Size(239, 20);
			this.textBoxDocumentCollection.TabIndex = 5;
			// 
			// buttonGenerateDocCollection
			// 
			this.buttonGenerateDocCollection.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
			this.buttonGenerateDocCollection.Cursor = System.Windows.Forms.Cursors.Hand;
			this.buttonGenerateDocCollection.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.buttonGenerateDocCollection.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.buttonGenerateDocCollection.ForeColor = System.Drawing.Color.RoyalBlue;
			this.buttonGenerateDocCollection.Image = ((System.Drawing.Image)(resources.GetObject("buttonGenerateDocCollection.Image")));
			this.buttonGenerateDocCollection.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.buttonGenerateDocCollection.Location = new System.Drawing.Point(10, 47);
			this.buttonGenerateDocCollection.Name = "buttonGenerateDocCollection";
			this.buttonGenerateDocCollection.Size = new System.Drawing.Size(381, 39);
			this.buttonGenerateDocCollection.TabIndex = 4;
			this.buttonGenerateDocCollection.Text = "Generate the entered Document Collection";
			this.buttonGenerateDocCollection.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.buttonGenerateDocCollection.UseVisualStyleBackColor = true;
			this.buttonGenerateDocCollection.Click += new System.EventHandler(this.buttonGenerateDocCollection_Click);
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
			this.statusStrip1.Location = new System.Drawing.Point(0, 595);
			this.statusStrip1.Name = "statusStrip1";
			this.statusStrip1.Size = new System.Drawing.Size(762, 22);
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
			this.groupBoxEmails.Location = new System.Drawing.Point(12, 336);
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
			// buttonTestHTML
			// 
			this.buttonTestHTML.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.buttonTestHTML.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.buttonTestHTML.Location = new System.Drawing.Point(541, 471);
			this.buttonTestHTML.Name = "buttonTestHTML";
			this.buttonTestHTML.Size = new System.Drawing.Size(206, 54);
			this.buttonTestHTML.TabIndex = 10;
			this.buttonTestHTML.Text = "Test HTML Agile Pack";
			this.buttonTestHTML.UseVisualStyleBackColor = true;
			this.buttonTestHTML.Click += new System.EventHandler(this.buttonTestHTML_Click);
			// 
			// buttonGenerateTestDoc
			// 
			this.buttonGenerateTestDoc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.buttonGenerateTestDoc.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.buttonGenerateTestDoc.Location = new System.Drawing.Point(541, 531);
			this.buttonGenerateTestDoc.Name = "buttonGenerateTestDoc";
			this.buttonGenerateTestDoc.Size = new System.Drawing.Size(206, 53);
			this.buttonGenerateTestDoc.TabIndex = 11;
			this.buttonGenerateTestDoc.Text = "Generate Word Document from Test HTML";
			this.buttonGenerateTestDoc.UseVisualStyleBackColor = true;
			this.buttonGenerateTestDoc.Click += new System.EventHandler(this.buttonGenerateTestDoc_Click);
			// 
			// buttonTestSomething
			// 
			this.buttonTestSomething.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.buttonTestSomething.Location = new System.Drawing.Point(541, 433);
			this.buttonTestSomething.Name = "buttonTestSomething";
			this.buttonTestSomething.Size = new System.Drawing.Size(205, 32);
			this.buttonTestSomething.TabIndex = 12;
			this.buttonTestSomething.Text = "Test Something";
			this.buttonTestSomething.UseVisualStyleBackColor = true;
			this.buttonTestSomething.Click += new System.EventHandler(this.buttonTestSomething_Click);
			// 
			// groupBoxPlatform
			// 
			this.groupBoxPlatform.Controls.Add(this.radioButtonDEV);
			this.groupBoxPlatform.Controls.Add(this.radioButtonQA);
			this.groupBoxPlatform.Controls.Add(this.radioButtonPROD);
			this.groupBoxPlatform.Cursor = System.Windows.Forms.Cursors.Hand;
			this.groupBoxPlatform.Location = new System.Drawing.Point(100, 57);
			this.groupBoxPlatform.Name = "groupBoxPlatform";
			this.groupBoxPlatform.Size = new System.Drawing.Size(200, 103);
			this.groupBoxPlatform.TabIndex = 0;
			this.groupBoxPlatform.TabStop = false;
			this.groupBoxPlatform.Text = "Platform";
			// 
			// radioButtonDEV
			// 
			this.radioButtonDEV.AutoSize = true;
			this.radioButtonDEV.Location = new System.Drawing.Point(18, 73);
			this.radioButtonDEV.Name = "radioButtonDEV";
			this.radioButtonDEV.Size = new System.Drawing.Size(88, 17);
			this.radioButtonDEV.TabIndex = 2;
			this.radioButtonDEV.TabStop = true;
			this.radioButtonDEV.Text = "Development";
			this.radioButtonDEV.UseVisualStyleBackColor = true;
			this.radioButtonDEV.CheckedChanged += new System.EventHandler(this.radioButtonDEV_CheckedChanged);
			// 
			// radioButtonQA
			// 
			this.radioButtonQA.AutoSize = true;
			this.radioButtonQA.Location = new System.Drawing.Point(18, 47);
			this.radioButtonQA.Name = "radioButtonQA";
			this.radioButtonQA.Size = new System.Drawing.Size(110, 17);
			this.radioButtonQA.TabIndex = 1;
			this.radioButtonQA.TabStop = true;
			this.radioButtonQA.Text = "Quality Assurance";
			this.radioButtonQA.UseVisualStyleBackColor = true;
			this.radioButtonQA.CheckedChanged += new System.EventHandler(this.radioButtonQA_CheckedChanged);
			// 
			// radioButtonPROD
			// 
			this.radioButtonPROD.AutoSize = true;
			this.radioButtonPROD.Location = new System.Drawing.Point(18, 21);
			this.radioButtonPROD.Name = "radioButtonPROD";
			this.radioButtonPROD.Size = new System.Drawing.Size(76, 17);
			this.radioButtonPROD.TabIndex = 0;
			this.radioButtonPROD.TabStop = true;
			this.radioButtonPROD.Text = "Production";
			this.radioButtonPROD.UseVisualStyleBackColor = true;
			this.radioButtonPROD.CheckedChanged += new System.EventHandler(this.radioButtonPROD_CheckedChanged);
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(762, 617);
			this.Controls.Add(this.groupBoxPlatform);
			this.Controls.Add(this.buttonTestSomething);
			this.Controls.Add(this.buttonGenerateTestDoc);
			this.Controls.Add(this.buttonTestHTML);
			this.Controls.Add(this.groupBoxEmails);
			this.Controls.Add(this.statusStrip1);
			this.Controls.Add(this.groupBoxSpecific);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.pictureBox1);
			this.Controls.Add(this.buttonGenerateAll);
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
			this.groupBoxPlatform.ResumeLayout(false);
			this.groupBoxPlatform.PerformLayout();
			this.ResumeLayout(false);

			}

		#endregion

		private System.Windows.Forms.Button buttonGenerateAll;
		private System.Windows.Forms.PictureBox pictureBox1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox groupBoxSpecific;
		private System.Windows.Forms.Button buttonGenerateDocCollection;
		private System.Windows.Forms.Label label2;
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
		private System.Windows.Forms.Button buttonTestHTML;
		private System.Windows.Forms.Button buttonGenerateTestDoc;
		private System.Windows.Forms.Button buttonTestSomething;
		private System.Windows.Forms.GroupBox groupBoxPlatform;
		private System.Windows.Forms.RadioButton radioButtonDEV;
		private System.Windows.Forms.RadioButton radioButtonQA;
		private System.Windows.Forms.RadioButton radioButtonPROD;
		private System.Windows.Forms.TextBox textBoxDocumentCollection;
		private System.Windows.Forms.ToolTip toolTip;
		}
	}

