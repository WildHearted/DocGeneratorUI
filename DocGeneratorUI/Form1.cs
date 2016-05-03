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

			MessageBox.Show("Here we go, ready to run a cycle of the DocGeneratorCore Controller as soon as you click Ok!", 
				"Ready to run the DocGenerator... message.", MessageBoxButtons.OK, MessageBoxIcon.Hand);

			// Declare the main controller object
			MainController objMainController = new MainController();
			objMainController.MainProcess();

			Cursor.Current = Cursors.Default;
			}
		}
	}
