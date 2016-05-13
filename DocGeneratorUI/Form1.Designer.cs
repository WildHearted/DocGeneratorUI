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
			this.progressBar1 = new System.Windows.Forms.ProgressBar();
			((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
			this.groupBoxSpecific.SuspendLayout();
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
			this.button1.Location = new System.Drawing.Point(100, 54);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(361, 39);
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
			this.label1.Size = new System.Drawing.Size(361, 42);
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
			this.groupBoxSpecific.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.groupBoxSpecific.Location = new System.Drawing.Point(12, 142);
			this.groupBoxSpecific.Name = "groupBoxSpecific";
			this.groupBoxSpecific.Size = new System.Drawing.Size(449, 77);
			this.groupBoxSpecific.TabIndex = 3;
			this.groupBoxSpecific.TabStop = false;
			this.groupBoxSpecific.Text = "Generate specific Document Collection";
			// 
			// maskedTextBox1
			// 
			this.maskedTextBox1.AsciiOnly = true;
			this.maskedTextBox1.BeepOnError = true;
			this.maskedTextBox1.Location = new System.Drawing.Point(135, 30);
			this.maskedTextBox1.Mask = "0000000";
			this.maskedTextBox1.Name = "maskedTextBox1";
			this.maskedTextBox1.Size = new System.Drawing.Size(58, 20);
			this.maskedTextBox1.TabIndex = 5;
			// 
			// button2
			// 
			this.button2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
			this.button2.Cursor = System.Windows.Forms.Cursors.Hand;
			this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.button2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(105)))), ((int)(((byte)(190)))), ((int)(((byte)(42)))));
			this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
			this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button2.Location = new System.Drawing.Point(258, 19);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(169, 39);
			this.button2.TabIndex = 4;
			this.button2.Text = "Generate";
			this.button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button2.UseVisualStyleBackColor = true;
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label2.Location = new System.Drawing.Point(7, 33);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(122, 13);
			this.label2.TabIndex = 0;
			this.label2.Text = "Document Collection ID:";
			// 
			// progressBar1
			// 
			this.progressBar1.Cursor = System.Windows.Forms.Cursors.No;
			this.progressBar1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(105)))), ((int)(((byte)(190)))), ((int)(((byte)(42)))));
			this.progressBar1.Location = new System.Drawing.Point(12, 225);
			this.progressBar1.MarqueeAnimationSpeed = 10;
			this.progressBar1.Name = "progressBar1";
			this.progressBar1.Size = new System.Drawing.Size(449, 23);
			this.progressBar1.Step = 5;
			this.progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
			this.progressBar1.TabIndex = 4;
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(470, 253);
			this.Controls.Add(this.progressBar1);
			this.Controls.Add(this.groupBoxSpecific);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.pictureBox1);
			this.Controls.Add(this.button1);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "Form1";
			this.Text = "DocGenerator User Interface";
			this.Load += new System.EventHandler(this.Form1_Load);
			((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
			this.groupBoxSpecific.ResumeLayout(false);
			this.groupBoxSpecific.PerformLayout();
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
		private System.Windows.Forms.ProgressBar progressBar1;
		}
	}

