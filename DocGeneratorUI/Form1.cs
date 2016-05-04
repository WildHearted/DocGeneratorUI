using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DocGeneratorCore;

namespace DocGeneratorUI
	{
	public partial class Form1:Form
		{
		public Form1()
			{
			InitializeComponent();
			}

		private void Form1_Load(object sender, EventArgs e)
			{

			}

		private void button1_Click(object sender, EventArgs e)
			{
			Cursor.Current = Cursors.WaitCursor;

			// Declare the main controller object
			MainController objMainController = new MainController();
			objMainController.MainProcess();

			Cursor.Current = Cursors.Default;
			MessageBox.Show("Successfully completed a cycle run of the DocGeneratorCore controller and processed all unprocessed "
				+ " Document Collections" + " \nClick it again to run another cycle.",
				"Run successfully completed.", MessageBoxButtons.OK, MessageBoxIcon.Information);

			}
		}
	}
