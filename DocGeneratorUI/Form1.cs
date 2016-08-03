using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Services.Client;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using HtmlAgilityPack;
using System.Runtime.Remoting.Contexts;
using System.Threading;
using System.Windows.Forms;
using DocGeneratorCore;

namespace DocGeneratorUI
	{
	public partial class Form1:Form
		{
		private CompleteDataSet completeDataSet;
		private static readonly Object lockThreadDatRefresh = new Object();

		public Form1()
			{
			Thread.CurrentThread.Name = "MainGUI";
			InitializeComponent();
			}

		private void Form1_Load(object sender, EventArgs e)
			{
			if(this.completeDataSet == null)
				{
				this.completeDataSet = new CompleteDataSet();
				this.statusStrip1.Text = "Dataset not cached yet.";
				this.statusStrip1.ForeColor = Color.Maroon;
				}
			
			Application.DoEvents();
			}

		private void button1_Click(object sender, EventArgs e)
			{

			if(this.comboBoxPlatform.SelectedItem == null 
				|| this.comboBoxPlatform.ToString().StartsWith("Select"))
				{
				MessageBox.Show(
					text: "Please specify a Platform before proceeding.",
					caption: "No Platform specified",
					buttons: MessageBoxButtons.OK,
					icon: MessageBoxIcon.Error);
				this.comboBoxPlatform.Focus();
				return;
				}

			Cursor.Current = Cursors.WaitCursor;


			String strExceptionMessage = String.Empty;
			// Initialise the listDocumentCollections object if it is null.
			List<DocumentCollection> listDocumentCollections = new List<DocumentCollection>();

			try
				{
				var dsDocCollections = 
					from dsDocumentCollection in this.completeDataSet.SDDPdatacontext.DocumentCollectionLibrary
					where dsDocumentCollection.GenerateActionValue != null
					&& dsDocumentCollection.GenerateActionValue != "Save but don't generate the documents yet"
					&& (dsDocumentCollection.GenerationStatus == enumGenerationStatus.Pending.ToString()
					|| dsDocumentCollection.GenerationStatus == null)
					orderby dsDocumentCollection.Modified
					select dsDocumentCollection;

				foreach(var recDocCollectionToGenerate in dsDocCollections)
					{
					// Create a DocumentCollection instance and populate the basic attributes.
					DocumentCollection objDocumentCollection = new DocumentCollection();
					objDocumentCollection.ID = recDocCollectionToGenerate.Id;
					if(recDocCollectionToGenerate.Title == null)
						objDocumentCollection.Title = "Collection Title for entry " + recDocCollectionToGenerate.Id;
					else
						objDocumentCollection.Title = recDocCollectionToGenerate.Title;
					objDocumentCollection.DetailComplete = false;
					// Add the DocumentCollection object to the listDocumentCollection
					listDocumentCollections.Add(objDocumentCollection);
					}

				// Check if there are any Document Collections to generate
				if(listDocumentCollections.Count > 0)
					{
					foreach(DocumentCollection entryDocumentCollection in listDocumentCollections)
						{// Invoke the DocGeneratorCore's MainController object MainProcess method
						MainController objMainController = new MainController();
						objMainController.DocumentCollectionsToGenerate = listDocumentCollections;
						objMainController.MainProcess(parDataSet: ref this.completeDataSet);
						}
					}

				MessageBox.Show("Successfully completed a cycle run of the DocGeneratorCore controller and processed all unprocessed "
				+ " Document Collections" + " \nClick it again to run another cycle.",
				"Run successfully completed.", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
			catch(DataServiceClientException exc)
				{
				strExceptionMessage = "*** Exception ERROR ***: DocGeneratorServer cannot access site: " 
					+ completeDataSet.SharePointSiteURL + completeDataSet.SharePointSiteSubURL
					+ " Please check that the computer/server is connected to the Domain network "
					+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult + "\nStatusCode: " + exc.StatusCode
					+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
				Cursor.Current = Cursors.Default;
				MessageBox.Show(strExceptionMessage,
					"Unable to generatate any documents.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
				//EmailBodyText += "\n\t - Unable to generatate any documents. \n" + strExceptionMessage;
				//bSuccessfulSentEmail = eMail.SendEmail(
				//	parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
				//	parSubject: "Error occurred in DocGenerator Server module.)",
				//	parBody: EmailBodyText,
				//	parSendBcc: false);
				}
			catch(DataServiceQueryException exc)
				{
				strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " 
					+ completeDataSet.SharePointSiteURL + completeDataSet.SharePointSiteSubURL
					+ " Please check that the computer/server is connected to the Domain network "
					+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
					+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
				Cursor.Current = Cursors.Default;
				MessageBox.Show(strExceptionMessage,
					"Unable to generatate any documents.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
				//EmailBodyText += "\n\t - Unable to generatate any documents. \n" + strExceptionMessage;
				//bSuccessfulSentEmail = eMail.SendEmail(
				//	parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
				//	parSubject: "Error occurred in DocGenerator Server module.)",
				//	parBody: EmailBodyText,
				//	parSendBcc: false);
				}
			catch(DataServiceRequestException exc)
				{
				strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " 
					+ completeDataSet.SharePointSiteURL + completeDataSet.SharePointSiteSubURL
					+ " Please check that the computer/server is connected to the Domain network "
					+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
					+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
				Cursor.Current = Cursors.Default;
				MessageBox.Show(strExceptionMessage,
					"Unable to generatate any documents.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
				//EmailBodyText += "\n\t - Unable to generatate any documents. \n" + strExceptionMessage;
				//bSuccessfulSentEmail = eMail.SendEmail(
				//	parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
				//	parSubject: "Error occurred in DocGenerator Server module.)",
				//	parBody: EmailBodyText,
				//	parSendBcc: false);
				}
			catch(DataServiceTransportException exc)
				{
				strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " 
					+ completeDataSet.SharePointSiteURL + completeDataSet.SharePointSiteSubURL
					+ " Please check that the computer/server is connected to the Domain network "
					+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
					+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
				Cursor.Current = Cursors.Default;
				MessageBox.Show(strExceptionMessage,
					"Unable to generatate any documents.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
				//EmailBodyText += "\n\t - Unable to generatate any documents. \n" + strExceptionMessage;
				//bSuccessfulSentEmail = eMail.SendEmail(
				//	parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
				//	parSubject: "Error occurred in DocGenerator Server module.)",
				//	parBody: EmailBodyText,
				//	parSendBcc: false);
				}
			catch(Exception exc)
				{
				if(exc.HResult == -2146330330)
					{
					strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " 
					+ completeDataSet.SharePointSiteURL + completeDataSet.SharePointSiteSubURL
					+ " Please check that the computer/server is connected to the Domain network "
					+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
					+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
					Cursor.Current = Cursors.Default;
					MessageBox.Show(strExceptionMessage,
						"Unable to generatate any documents.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					return;
					//EmailBodyText += "\n\t - Unable to generatate any documents. \n" + strExceptionMessage;
					//bSuccessfulSentEmail = eMail.SendEmail(
					//	parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
					//	parSubject: "Error occurred in DocGenerator Server module.)",
					//	parBody: EmailBodyText,
					//	parSendBcc: false);
					}
				else if(exc.HResult == -2146233033)
					{
					strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " 
						+ completeDataSet.SharePointSiteURL + completeDataSet.SharePointSiteSubURL
						+ " Please check that the computer/server is connected to the Domain network "
						+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
						+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
					Cursor.Current = Cursors.Default;
					MessageBox.Show(strExceptionMessage,
						"Unable to generatate any documents.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					return;
					//EmailBodyText += "\n\t - Unable to generatate any documents. \n" + strExceptionMessage;
					//bSuccessfulSentEmail = eMail.SendEmail(
					//	parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
					//	parSubject: "Error occurred in DocGenerator Server module.)",
					//	parBody: EmailBodyText,
					//	parSendBcc: false);
					}
				else
					{
					strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " 
						+ completeDataSet.SharePointSiteURL + completeDataSet.SharePointSiteSubURL
						+ " Please check that the computer/server is connected to the Domain network "
						+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
						+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
					Cursor.Current = Cursors.Default;
					MessageBox.Show(strExceptionMessage,
						"Unable to generatate any documents.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					return;
					//EmailBodyText += "\n\t - Unable to generatate any documents. \n" + strExceptionMessage;
					//bSuccessfulSentEmail = eMail.SendEmail(
					//	parRecipient: Properties.Resources.EmailAddress_TechnicalSupport,
					//	parSubject: "Error occurred in DocGenerator Server module.)",
					//	parBody: EmailBodyText,
					//	parSendBcc: false);
					};
				}
			finally
				{
				Cursor.Current = Cursors.Default;
				}
			
			}

//===B
		private void button2_Click(object sender, EventArgs e)
			{
			// Validate the input

			if(this.comboBoxPlatform.SelectedItem == null
				|| this.comboBoxPlatform.ToString().StartsWith("Select"))
				{
				MessageBox.Show(
					text: "Please specify a Platform before proceeding.",
					caption: "No Platform specified",
					buttons: MessageBoxButtons.OK,
					icon: MessageBoxIcon.Error);
				this.comboBoxPlatform.Focus();
				return;
				}

			if(maskedTextBox1.Text == null || maskedTextBox1.Text == "")
				{
				MessageBox.Show("Please enter a numeric value, before clicking the Generate button",
						"No value entred.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
				}
			int intDocumentCollectionID;
			bool bIsNumeric = int.TryParse(maskedTextBox1.Text.Trim(), out intDocumentCollectionID);
			if(!bIsNumeric)
				{
				MessageBox.Show("Only numeric values can be entered.",
						"Invalid value entered.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
				}
			// -----------------------------------
			Cursor.Current = Cursors.WaitCursor;
			if(this.completeDataSet.IsDataSetComplete)
				{
				this.toolStripStatusLabel.Text = "Dataset cached and ready.";
				this.toolStripStatusLabel.ForeColor = Color.Green;
				}
			else
				{
				this.toolStripStatusLabel.Text = "Dataset is not cached, will be loaded now...";
				this.toolStripStatusLabel.ForeColor = Color.Maroon;
				}

			Application.DoEvents();

			string strExceptionMessage = String.Empty;
			// Initialise the listDocumentCollections object if it is null.
			List<DocumentCollection> listDocumentCollections = new List<DocumentCollection>();

			try
				{
				var dsDocumentCollections = 
					from dsDC in this.completeDataSet.SDDPdatacontext.DocumentCollectionLibrary
					where dsDC.Id == intDocumentCollectionID
					select dsDC;

				var objDocCollection = dsDocumentCollections.FirstOrDefault();

				if(objDocCollection == null)
					{
					Cursor.Current = Cursors.Default;
					MessageBox.Show("The Document Collection ID that you entered doesn't exist. Please enter a valid ID.",
						"Document Collection ID, doesn't exist.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					return;
					}
				else
					{
					// Create a DocumentCollection instance and populate the basic attributes.
					Application.DoEvents();
					DocumentCollection objDocumentCollection = new DocumentCollection();
					objDocumentCollection.ID = objDocCollection.Id;
					if(objDocCollection.Title == null)
						objDocumentCollection.Title = "Collection Title for entry " + objDocCollection.Id;
					else
						objDocumentCollection.Title = objDocCollection.Title;

					objDocumentCollection.DetailComplete = false;
					// Add the DocumentCollection object to the listDocumentCollection
					listDocumentCollections.Add(objDocumentCollection);
					}
				Application.DoEvents();

				Console.WriteLine("Is the DataSet complete? [{0}]", this.completeDataSet.IsDataSetComplete);
				
				// Check if there are any Document Collections to generate
				if(listDocumentCollections.Count > 0)
					{
					// Invoke the DocGeneratorCore's MainController object MainProcess method and send all the entries for processing
					MainController objMainController = new MainController();
					objMainController.DocumentCollectionsToGenerate = listDocumentCollections;
					objMainController.MainProcess(parDataSet: ref this.completeDataSet);
					}
				}
			catch(DataServiceClientException exc)
				{
				strExceptionMessage = "*** Exception ERROR ***: DocGeneratorUI cannot access site: " 
					+ completeDataSet.SharePointSiteURL + completeDataSet.SharePointSiteSubURL
					+ " Please check that the computer/server is connected to the Domain network "
					+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult + "\nStatusCode: " + exc.StatusCode
					+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
				Application.DoEvents();
				Cursor.Current = Cursors.Default;
				MessageBox.Show(strExceptionMessage,
					"Unable to generatate any documents.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
				}
			catch(DataServiceQueryException exc)
				{
				Cursor.Current = Cursors.Default;
				if(exc.InnerException.Message.Contains("Unauthorized"))
					{
					MessageBox.Show("An Authenication error occurred for Account: " 
						+ Properties.Resources.DocGenerator_AccountName 
						+ " is not authorised to access: " 
						+ completeDataSet.SharePointSiteURL + completeDataSet.SharePointSiteSubURL
						+ "\nPlease check if the password has not expired." ,
						"Unauthorised.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					}
				else
					{
					MessageBox.Show("First check that you are connected to the Network Domain, if not, establish VPN or LAN connection. Else, the Document Collection "
					+ "ID that you entered doesn't exist.",
						"Check Domain connection and that Document Collection ID exist.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					}
				return;
				}
			catch(DataServiceRequestException exc)
				{
				strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " 
					+ completeDataSet.SharePointSiteURL + completeDataSet.SharePointSiteSubURL
					+ " Please check that the computer/server is connected to the Domain network "
					+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
					+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
				Application.DoEvents();
				Cursor.Current = Cursors.Default;
				MessageBox.Show(strExceptionMessage,
					"Unable to generatate any documents.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
				}
			catch(DataServiceTransportException exc)
				{
				strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " 
					+ completeDataSet.SharePointSiteURL + completeDataSet.SharePointSiteSubURL
					+ " Please check that the computer/server is connected to the Domain network "
					+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
					+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
				Application.DoEvents();
				Cursor.Current = Cursors.Default;
				MessageBox.Show(strExceptionMessage,
					"Unable to generatate any documents.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
				}
			catch(Exception exc)
				{
				if(exc.HResult == -2146330330)
					{
					strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " 
						+ completeDataSet.SharePointSiteURL + completeDataSet.SharePointSiteSubURL
						+ " Please check that the computer/server is connected to the Domain network "
						+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
						+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
					Application.DoEvents();
					Cursor.Current = Cursors.Default;
					MessageBox.Show(strExceptionMessage,
						"Unable to generatate any documents.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					return;
					}
				else if(exc.HResult == -2146233033)
					{
					strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " 
						+ completeDataSet.SharePointSiteURL + completeDataSet.SharePointSiteSubURL
						+ " Please check that the computer/server is connected to the Domain network "
						+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
						+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
					Application.DoEvents();
					Cursor.Current = Cursors.Default;
					MessageBox.Show(strExceptionMessage,
						"Unable to generatate any documents.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					return;
					}
				else
					{
					strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " 
						+ completeDataSet.SharePointSiteURL + completeDataSet.SharePointSiteSubURL
						+ " Please check that the computer/server is connected to the Domain network "
						+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
						+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
					Application.DoEvents();
					Cursor.Current = Cursors.Default;
					MessageBox.Show(strExceptionMessage,
						"Unable to generatate any documents.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					return;
					};
				}
			finally
				{
				Cursor.Current = Cursors.Default;
				}

			if(this.completeDataSet.IsDataSetComplete)
				{
				this.statusStrip1.Text = "Dataset cached and ready.";
				this.statusStrip1.ForeColor = Color.Green;
				}
			else
				{
				this.statusStrip1.Text = "Dataset is not cached.";
				this.statusStrip1.ForeColor = Color.Maroon;
				}

			MessageBox.Show("Successfully completed the generation of Document Collection: " + maskedTextBox1.Text
				+ " \nClick it again to generate the same Document Collection or enter another number...",
				"Generation successfully completed.", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}

		/// <summary>
		///  This timer is fired every 60 seconds, to check if something changed in the dataset and to 
		///  refresh it if any data changed...
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void timerCheckDataset_Tick(object sender, EventArgs e)
			{
			Console.WriteLine("TTTT - TimerEvent: timerCheckDataSet");
			Thread.CurrentThread.Name = "PopulateDataset";
			DataCache datacashDataSet = new DataCache();
			//- Define and launch a separate thread that will begin to opulate the dataset.
			Thread threadPopulateDataset = new Thread(() => datacashDataSet.Populate_DataSet(ref completeDataSet));
			threadPopulateDataset.Name = "PopulateDataset";
			threadPopulateDataset.Start();
			}

		private void timerCheckDataSetStatus_Tick(object sender, EventArgs e)
			{
			if(this.completeDataSet == null)
				{
				this.toolStripStatusLabel.Text = "DataSet Cache not loaded. ";
				this.toolStripStatusLabel.ForeColor = Color.Red;
				Application.DoEvents();
				}
			else
				{
				if(this.completeDataSet.IsDataSetComplete == false)
					{
					this.toolStripStatusLabel.Text = "DataSet Cache incomplete.. ";
					this.toolStripStatusLabel.ForeColor = Color.Orange;
					Application.DoEvents();
					}
				else
					{
					this.toolStripStatusLabel.Text = "Dataset Cashe Loaded.";
					this.toolStripStatusLabel.ForeColor = Color.Green;
					Application.DoEvents();
					}
				}
			}

		private void buttonSendEmail_Click(object sender, EventArgs e)
			{
			if(this.comboBoxPlatform.SelectedItem == null
				|| this.comboBoxPlatform.ToString().StartsWith("Select"))
				{
				MessageBox.Show(
					text: "Please specify a Platform before proceeding.",
					caption: "No Platform specified",
					buttons: MessageBoxButtons.OK,
					icon: MessageBoxIcon.Error);
				this.comboBoxPlatform.Focus();
				return;
				}

			// Validate the e-mail input
			if(textBoxEmail.Text == null || textBoxEmail.Text == "")
				{
				MessageBox.Show("Please enter an email address before clicking the Send Email button",
						"No Email adddress entered.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
				}

			if(!textBoxEmail.Text.Contains("@"))
				{
				MessageBox.Show("Please enter a VALID email address before clicking the Send Email button",
						"Invalid Email adddress entered.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
				}
			string EmailBodyText = String.Empty;

			//+ Define an User Email message content
			Cursor.Current = Cursors.WaitCursor;
			EmailModel emailModel = new EmailModel();
			emailModel.Name = "Ben";
			emailModel.EmailAddress = "ben.vandenberg@dimensiondata.com";
			emailModel.CollectionID = 1387;
			emailModel.CollectionTitle = "MSEN 2.0 in the Document Collections Library";
			emailModel.CollectionURL = "\"" + "https://teams.dimensiondata.com/sites/ServiceCatalogue/Lists/Document%20Collection%20Library/EditFormAJS.aspx?ID=1387" + "\"";
			emailModel.EmailGeneratedDocs = new List<EmailGeneratedDocuments>();
			
			//- first document was NOT successful - it had content errors
				EmailGeneratedDocuments objGeneratedDocument1 = new EmailGeneratedDocuments();
				objGeneratedDocument1.Title = "Internal Service Description Document with Deliverables, Reports, Meetings Inline";
				objGeneratedDocument1.URL = "https://teams.dimensiondata.com/sites/ServiceCatalogue/Generated%20Documents/ISD_Document_DRM_Inline_5-19-2016_8-37-40_AM.docx";
				objGeneratedDocument1.IsSuccessful = false;
				objGeneratedDocument1.Errors = new List<string>();
				objGeneratedDocument1.Errors.Add("Content Error: The Service Product ID: 177 contains an error in the Enhance Rich Text column ISD Description.The width of the TABLE is NOT defined.The table width must be set to a percentage(%) value in order for the DocGenerator to scale and fit in the document page width.");
				objGeneratedDocument1.Errors.Add("Content Error: Deliverable ID 86 contains an error in the Enhance Rich Text Input column / field.The width of a TABLE is NOT defined.The table width must be set to a percentage(%) value in order for the DocGenerator to scale and fit in the document page width.");
				//- Add the GeneratedDocument to the DocumentCollection's list of GeneratedDocs
				emailModel.EmailGeneratedDocs.Add(objGeneratedDocument1);

			//- first document was NOT successful - it had content errors
				EmailGeneratedDocuments objGeneratedDocument2 = new EmailGeneratedDocuments();
				objGeneratedDocument2.Title = "Internal Service Description Document with Deliverables, Reports, Meetings Sections";
				objGeneratedDocument2.URL = "https://teams.dimensiondata.com/sites/ServiceCatalogue/Generated%20Documents/ISD_Document_DRM_Sections_5-19-2016_8-37-40_AM.docx";
				objGeneratedDocument2.IsSuccessful = true;
				//- Add the GeneratedDocument to the DocumentCollection's list of GeneratedDocs
				emailModel.EmailGeneratedDocs.Add(objGeneratedDocument2);

			// Declare the Email object and assign the above defined message to the relevant property
			eMail objUserSuccessEmail = new eMail();
			objUserSuccessEmail.ConfirmationEmailModel = emailModel;
			//- Compile the HTML email message
			if(objUserSuccessEmail.ComposeHTMLemail(enumEmailType.UserSuccessfulConfirmation))
				{
				bool bSuccessfulSentEmail = objUserSuccessEmail.SendEmail(
					parDataSet: ref completeDataSet,
					parRecipient: this.textBoxEmail.Text,
					parSubject: "SDDP: End User Document Generation Confirmation Message (Sample)",
					parSendBcc: false);

				if(bSuccessfulSentEmail)
					{
					Cursor.Current = Cursors.Default;
					MessageBox.Show("The test HTML message was sent successfully...",
					"Email successfully sent.", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
				else
					{
					Cursor.Current = Cursors.Default;
					MessageBox.Show("Sending the e-mail failed. \nPlease check that you are connected to the Dimension Data Domain before you try again.",
					"The e-mail couldn't be send", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					}
				}
			else
				{
				Cursor.Current = Cursors.Default;
				MessageBox.Show("The application could not compile the HTML e-mail, please report it to the SDDP System Administator.",
					"The HTML e-mail compile FAILED.", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}

			}

		private void buttonSendTechnicalEmail_Click(object sender, EventArgs e)
			{
			if(this.comboBoxPlatform.SelectedItem == null
				|| this.comboBoxPlatform.ToString().StartsWith("Select"))
				{
				MessageBox.Show(
					text: "Please specify a Platform before proceeding.",
					caption: "No Platform specified",
					buttons: MessageBoxButtons.OK,
					icon: MessageBoxIcon.Error);
				this.comboBoxPlatform.Focus();
				return;
				}

			// Validate the e-mail input
			if(textBoxEmail.Text == null || textBoxEmail.Text == "")
				{
				MessageBox.Show("Please enter an email address before clicking the Send Email button",
						"No Email adddress entered.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
				}

			if(!textBoxEmail.Text.Contains("@"))
				{
				MessageBox.Show("Please enter a VALID email address before clicking the Send Email button",
						"Invalid Email adddress entered.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
				}
			string EmailBodyText = String.Empty;

			//+ Define an User Email message content
			Cursor.Current = Cursors.WaitCursor;
			TechnicalSupportModel emailModel = new TechnicalSupportModel();
			emailModel.EmailAddress = this.textBoxEmail.Text;
			emailModel.MessageLines = new List<string>();
			if(this.radioInformation.Checked)
				{
				emailModel.Classification = enumMessageClassification.Information;
				emailModel.MessageHeading = "For Your Information...";
				emailModel.Instruction = "This is a just For Your Information...";
				emailModel.MessageLines.Add("The DocGenerator Service was started at " + DateTime.Now.ToString());
				emailModel.MessageLines.Add("No need to worry...");
				}
			else if(radioWarning.Checked)
				{
				emailModel.Classification = enumMessageClassification.Warning;
				emailModel.MessageHeading = "Warning Message...";
				emailModel.Instruction = "This is a Warning message to inform you that...";
				emailModel.MessageLines.Add("The DocGenerator Service was stopped at " + DateTime.Now.ToString());
				emailModel.MessageLines.Add("Please lookout for the start message, if it doesn't appear soon, please investigate and restart the DocGenerator Service.");
				}
			else
				{
				emailModel.Classification = enumMessageClassification.Error;
				emailModel.MessageHeading = "The following unexpected error occurred in the DocGenerator:";
				emailModel.Instruction = "This message is to inform you that the an DocGenerator ERROR occurred...";
				emailModel.MessageLines.Add("An unexpected error occurred at " + DateTime.Now.ToString());
				emailModel.MessageLines.Add("Please investigate and resolve the error as soon as possible.");
				}

			// Declare the Email object and assign the above defined message to the relevant property
			eMail objTechnicalEmail = new eMail();
			objTechnicalEmail.TechnicalEmailModel = emailModel;
			//- Compile the HTML email message
			if(objTechnicalEmail.ComposeHTMLemail(enumEmailType.TechnicalSupport))
				{
				bool bSuccessfulSentEmail = objTechnicalEmail.SendEmail(
					parDataSet: ref completeDataSet,
					parRecipient: this.textBoxEmail.Text,
					parSubject: "SDDP: Technical Support email (Sample)",
					parSendBcc: false);

				if(bSuccessfulSentEmail)
					{
					Cursor.Current = Cursors.Default;
					MessageBox.Show("The test HTML message was sent successfully...",
					"Email successfully sent.", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
				else
					{
					Cursor.Current = Cursors.Default;
					MessageBox.Show("Sending the e-mail failed. \nPlease check that you are connected to the Dimension Data Domain before you try again.",
					"The e-mail couldn't be send", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					}
				}
			else
				{
				Cursor.Current = Cursors.Default;
				MessageBox.Show("The application could not compile the HTML e-mail, please report it to the SDDP System Administator.",
					"The HTML e-mail compile FAILED.", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
			}

		private void buttonSendUserErrorEmail_Click(object sender, EventArgs e)
			{
			if(this.comboBoxPlatform.SelectedItem == null
				|| this.comboBoxPlatform.ToString().StartsWith("Select"))
				{
				MessageBox.Show(
					text: "Please specify a Platform before proceeding.",
					caption: "No Platform specified",
					buttons: MessageBoxButtons.OK,
					icon: MessageBoxIcon.Error);
				this.comboBoxPlatform.Focus();
				return;
				}

			// Validate the e-mail input
			if(textBoxEmail.Text == null || textBoxEmail.Text == "")
				{
				MessageBox.Show("Please enter an email address before clicking the Send Email button",
						"No Email adddress entered.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
				}

			if(!textBoxEmail.Text.Contains("@"))
				{
				MessageBox.Show("Please enter a VALID email address before clicking the Send Email button",
						"Invalid Email adddress entered.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
				}
			string EmailBodyText = String.Empty;

			//+ Define an User Error message content
			Cursor.Current = Cursors.WaitCursor;
			EmailModel emailModel = new EmailModel();
			emailModel.Name = "Ben";
			emailModel.EmailAddress = "ben.vandenberg@dimensiondata.com";
			emailModel.CollectionID = 1387;
			emailModel.CollectionTitle = "MSEN 2.0 in the Document Collections Library";
			emailModel.CollectionURL = "\"" + "https://teams.dimensiondata.com/sites/ServiceCatalogue/Lists/Document%20Collection%20Library/EditFormAJS.aspx?ID=1387" + "\"";
			emailModel.Failed = true;
			emailModel.Error = "Unfortunatley, you submitted the Document Collection without specifing any document(s) to be generated.";

			// Declare the Email object and assign the above defined message to the relevant property
			eMail objUserSuccessEmail = new eMail();
			objUserSuccessEmail.ConfirmationEmailModel = emailModel;
			//- Compile the HTML email message
			if(objUserSuccessEmail.ComposeHTMLemail(enumEmailType.UserErrorConfirmation))
				{
				bool bSuccessfulSentEmail = objUserSuccessEmail.SendEmail(
					parDataSet: ref completeDataSet,
					parRecipient: this.textBoxEmail.Text,
					parSubject: "SDDP: End User Document Generation Confirmation Error Message (Sample)",
					parSendBcc: false);

				if(bSuccessfulSentEmail)
					{
					Cursor.Current = Cursors.Default;
					MessageBox.Show("The test HTML message was sent successfully...",
					"Email successfully sent.", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
				else
					{
					Cursor.Current = Cursors.Default;
					MessageBox.Show("Sending the e-mail failed. \nPlease check that you are connected to the Dimension Data Domain before you try again.",
					"The e-mail couldn't be send", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					}
				}
			else
				{
				Cursor.Current = Cursors.Default;
				MessageBox.Show("The application could not compile the HTML e-mail, please report it to the SDDP System Administator.",
					"The HTML e-mail compile FAILED.", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
			}

		private void comboBoxPlatform_SelectedIndexChanged(object sender, EventArgs e)
			{
			if(this.comboBoxPlatform.SelectedItem.ToString() == "Development Environment")
				{//- Development Environment
				completeDataSet.SharePointSiteURL = Properties.Resources.SharePointURL_Dev;
				completeDataSet.SharePointSiteSubURL = Properties.Resources.SharePointSiteURL_Dev;
				completeDataSet.IsDataSetComplete = false;
				}
			else if(this.comboBoxPlatform.SelectedItem.ToString() == "Quality Assurance Environment")
				{//- QA Environment
				completeDataSet.SharePointSiteURL = Properties.Resources.SharePointURL_QA;
				completeDataSet.SharePointSiteSubURL = Properties.Resources.SharePointSiteURL_QA;
				completeDataSet.IsDataSetComplete = false;
				}
			else
				{//- Production Environment
				completeDataSet.SharePointSiteURL = Properties.Resources.SharePointURL_Prod;
				completeDataSet.SharePointSiteSubURL = Properties.Resources.SharePointSiteURL_Prod;
				completeDataSet.IsDataSetComplete = false;
				}

			//- Initiate the SharePoint Datacontext...
			this.completeDataSet.SDDPdatacontext = new DocGeneratorCore.SDDPServiceReference.DesignAndDeliveryPortfolioDataContext(
					new Uri(completeDataSet.SharePointSiteURL + completeDataSet.SharePointSiteSubURL + Properties.Resources.SharePointRESTuri));

			this.completeDataSet.SDDPdatacontext.Credentials = new NetworkCredential(
				userName: Properties.Resources.DocGenerator_AccountName,
				password: Properties.Resources.DocGenerator_Account_Password,
				domain: Properties.Resources.DocGenerator_AccountDomain);

			this.completeDataSet.SDDPdatacontext.MergeOption = MergeOption.NoTracking;

			}

		private void buttonTestHTML_Click(Object sender, EventArgs e)
			{
			string htmlText =
				"<p> This section validates the<em> cascading numbered bullets</ em >.</ p ><lo><li> This is the 1st for Level 1 </li><li> This is the 2nd for Level 1 </li><li> This is the 3rd for Level 1 </li><lo><li> This is the 1st for Level 2 </li><li> This is the 2nd for Level 2 </li><lo><li> This is the 1st for Level 3 </li><lo><li> This is the 1st for Level 4 </li><lo><li> This is the 1st for Level 5 </li></ol></ol><li> This is the 2nd for Level 3 </li></ol><li> is the 3rd for Level 2 </li><li> This is the 4th for Level 2 </li></ol><li> This is the 4th for Level 1 </li></ol><p> This conclude the numbered list testing </p>";

			//-Declare the htmlData object instance
			HtmlAgilityPack.HtmlDocument htmlData = new HtmlAgilityPack.HtmlDocument();
			//-Load the HTML content into the htmlData object.
			//htmlData.Load(@"E:\Development\Projects\DocGeneratorGUI\DocGeneratorUI\TestData\HTMLtestPage_CascadedBulletList.html");
			htmlData.Load(@"E:\Development\Projects\DocGeneratorGUI\DocGeneratorUI\TestData\HTMLtestPage_Everything.html");

			//-Load a string into the HTMLDocument
			Console.WriteLine("______________________________HTML starts______________________________________");
			//htmlData.LoadHtml(htmlText);
			Console.WriteLine("HTML: {0}", htmlData.DocumentNode.OuterHtml);
			Console.WriteLine("_______________________________________________________________________________");
			//- Set the ROOT of the data loaded in the htmlData
			var htmlRoot = htmlData.DocumentNode;
			Console.WriteLine("Root Node Tag..............: {0}", htmlRoot.Name);

			/*Console.WriteLine("\t - Has Child Nodes......? {0}", htmlRoot.HasChildNodes);
			*/

			//- The sequence of the values in the tuple variable is: 
			//-**1** *Bullet Levels*, 
			//-**2** *Number Levels* 
			Tuple<int, int> headingbulletNumberLevels = new Tuple<int, int>(0, 0);
			int bulletLevel = 0;
			int numberLevel = 0;
			int headingLevel = 0;
			bool isBold = false;
			bool isItalics = false;
			bool isUnderline = false;
			bool isSubscript = false;
			bool isSuperscript = false;
			string clientName = "Software Engineering Services";
			string cleanText = String.Empty;
			string captionText = String.Empty;
			string tableAttrValue = string.Empty;	//-Table Attribute value **work string** 
			int tableWidth = 0;
			int columnWidth = 0;
			bool isHeaderRow = false;			//-indicated whether the current table row is **HeaderRow**
			int tableRowSpanQty = 0;				//-variable indicates the **number** of table *rows* to span
			int tableColumnSpanQty = 0;			//-variable indicating the **number** of table *columns* to span
			foreach(HtmlNode node in htmlData.DocumentNode.Descendants())
				{
				
				switch(node.Name)
					{
					case "div":
						{
						//-Check for other tags
						if(node.HasChildNodes)
							{
							//- Just ignore the div tag
							}
						else
							{
							cleanText = CleanText(node.InnerText, clientName);
							if(cleanText != String.Empty)
								Console.Write(" <{0}>|{1}|", node.Name, cleanText);
							}
						break;
						}

					case "#text":
						{
						if(node.HasChildNodes)
							{
							//!to do!
							}
						else
							{
							cleanText = CleanText(node.InnerText, clientName);
							if(cleanText != String.Empty)
								{
								headingLevel = 0;
								if(node.XPath.Contains("/h1"))
									headingLevel = 1;
								if(node.XPath.Contains("/h2"))
									headingLevel = 2;
								if(node.XPath.Contains("/h3"))
									headingLevel = 3;
								if(node.XPath.Contains("/h4)"))
									headingLevel = 4;


								if(headingLevel > 0)
									{
									Console.Write("\n {0} + <h{1}>", Tabs(headingLevel), headingLevel);
									}

								if(node.XPath.Contains("/strong"))
									Console.Write("|BOLD|");
								if(node.XPath.Contains("/em"))
									Console.Write("|ITALICS|");
								if(node.XPath.Contains("/span"))
									{
									Console.Write("|UNDELINE|");
									}
								if(node.XPath.Contains("/sub"))
									Console.Write("|SUBSCRIPT|");
								if(node.XPath.Contains("/sup"))
									Console.Write("|SUPERSCRIPT|");
								if(cleanText != String.Empty)
									Console.Write("[{0}]", cleanText);
								}
							}
						break;
						}

					case "p":
						{
						//-Get the number of bullet - and number - levles in the xPath
						headingbulletNumberLevels = GetBulletNumberLevels(node.XPath);
						headingLevel = headingbulletNumberLevels.Item1;
						bulletLevel = headingbulletNumberLevels.Item2;

						//-Check if the paragraph is part of a bullet- number- list
						if(numberLevel > 0 || bulletLevel > 0)
							{
							//-don't insert a paragraph, because the text is part of the bullet or number
							break;
							}

						if(node.HasChildNodes)
							{
							Console.Write("\n <{0}> ", node.Name);
							}
						else
							{
							cleanText = node.InnerText;
							if(cleanText != String.Empty)
								Console.Write(" <{0}>", node.Name);
							}
						break;
						}
					
					//+Heading 1-4
					/*case "h1":
					case "h2":
					case "h3":
					case "h4":
						{
						//- Set **Heading** level
						headingLevel = Convert.ToInt16(node.Name.Substring(1, 1));
						if(node.HasChildNodes)
							{
							Console.Write("\n {0} + <{1}>", Tabs(headingLevel + bulletLevel), node.Name);
							}
						else
							{
							cleanText = node.InnerText;
							Console.Write("\t\t\t <{0}>|{1}|", node.Name, cleanText);
							}
						break;
						}*/
					//+ Unorganised List - **<ul>**
					case "ul":
						{
						if(node.HasChildNodes)
							{
							//Console.Write("\n {0} <{1}>", Tabs(headingLevel) + Tabs(bulletLevel), node.Name);
							}
						else
							{
							cleanText = node.InnerText;
							Console.WriteLine("\t\t\t <{0}>|{1}|", node.Name, cleanText);
							}
						break;
						}

					//+ Organised List - **<ol>**
					case "ol":
						{
						if(node.HasChildNodes)
							{
							//Console.Write("\n {0} <{1}>", Tabs(headingLevel) + Tabs(bulletLevel), node.Name);
							}
						else
							{
							cleanText = node.InnerText;
							Console.WriteLine("\t\t\t <{0}>|{1}|", node.Name, cleanText);
							}
						break;
						}

					//+List Item - **<li>**
					case "li":
						{
						//-Get the number of bullet- and number- levles in the xPath
						headingbulletNumberLevels = GetBulletNumberLevels(node.XPath);
						headingLevel = headingbulletNumberLevels.Item1;
						bulletLevel = headingbulletNumberLevels.Item2;
						
						if(node.HasChildNodes)
							{
							if(bulletLevel > 0)
								{
								Console.Write("\n {0} - <{1}>", Tabs(headingLevel + bulletLevel), node.Name);
								}
							else if(numberLevel > 0)
								{
								Console.Write("\n {0} {1} <{2}>", Tabs(headingLevel  + numberLevel), numberLevel , node.Name);
								}
							}
						else
							{
							cleanText = node.InnerText;
							Console.WriteLine("\t\t\t <{0}>|{1}|", node.Name, cleanText);
							}
						break;
						}

					//++Image
					case "img":
						{


						break;
						}

					//++Table
					case "table":
						{
						Console.Write("\n\n <Table> ");
						tableWidth = 0;
						if(!node.HasAttributes)  //- The table doesn't have attributes 
							Console.Write("\n ERROR - No attributes defined for the table");
						else //-Process the table attributes
							{
							foreach(HtmlAgilityPack.HtmlAttribute tableAttr in node.Attributes)
								{
								switch(tableAttr.Name)
									{
									//-use the **summary** to obtain the **Table Caption**
									case "summary": //- get the table caption
										{
										if(tableAttr.Value == null)
											captionText = String.Empty;
										else
											captionText = tableAttr.Value;
										break;
										}

									//-get the table **width** from the style attribute
									case "style":
										{
										//- Check that the style contains the table width as part of the style
										if(tableAttr.Value.Contains("width:"))
											{
											if(tableAttr.Value.Contains("%"))
												{
												if(int.TryParse(tableAttr.Value.Substring(
													tableAttr.Value.IndexOf(":") + 2,
													(tableAttr.Value.IndexOf("%") - tableAttr.Value.IndexOf(":") - 2)),
													out tableWidth) == false)
													{
													//- Could not parse the integer which means the tableWidth remains as it was before the parse.
													}
												}
											else //-Table width is **NOT** a percentage :. px value
												{
												if(int.TryParse(tableAttr.Value.Substring(
													tableAttr.Value.IndexOf(":") + 2,
													(tableAttr.Value.IndexOf("px") - tableAttr.Value.IndexOf(":") - 2)),
													out tableWidth) == false)
													{
													//- Could not parse the integer which means the tableWidth remains as it was before the parse.
													}
												}
											}
										break;
										}

									//-The table Width is specified as an attribute
									case "width":
										{
										if(tableAttr.Value.Contains("%"))
											{
											if(int.TryParse(tableAttr.Value.Substring(
												tableAttr.Value.IndexOf(":") + 2,
												(tableAttr.Value.IndexOf("%") - tableAttr.Value.IndexOf(":") - 2)),
												out tableWidth) == false)
												{
												//- Could not parse the integer which means the tableWidth remains as it was before the parse.
												}
											}
										else //-Table width is **NOT** a percentage :. px value
											{
											if(int.TryParse(tableAttr.Value.Substring(
												tableAttr.Value.IndexOf(":") + 2,
												(tableAttr.Value.IndexOf("px") - tableAttr.Value.IndexOf(":") - 2)),
												out tableWidth) == false)
												{
												//- Could not parse the integer which means the tableWidth remains as it was before the parse.
												}
											}

										break;
										}


									}
								}


							}
						Console.Write("\t Width: {0}", tableWidth);
						break;
						}
					//+Table Body = **<tb>**
					case "tbody":
						{
						//-Just ignore it.
						break;
						}


					//+Table Row = **<tr>**
					case "tr":
						{
						Console.Write("\n\t <Row>");

						//Determine if it is a **HEADER** row
						if(!node.HasAttributes)  //- The table doesn't have attributes 
							{ //-Therefore it will also not have a Header Row
							isHeaderRow = false;
							}
						else //-Process the table attributes
							{

							foreach(HtmlAgilityPack.HtmlAttribute tableAttr in node.Attributes)
								{
								//use the **class** to dertermine whether the row is specified as a *TableHeader*
								if(tableAttr.Name == "class")
									{
									if(tableAttr.Value == null)
										isHeaderRow = false;
									else
										{
										if(tableAttr.Value.Contains("HeaderRow"))
											isHeaderRow = true;
										else
											isHeaderRow = false;
										}
									}
								}
							}
						Console.Write("Header Row: {0}", isHeaderRow);
						break;
						}

					//+Table Header Cell - **<th>**
					case "th":
						{
						//-If the **table header** was not determined in the class of the table row *<tr>* tag, then the table has a header row if a *<th>* tag is present
						isHeaderRow = true;
						//Determined the table has a **HEADER** row

						if(node.HasAttributes)  //- The table Header Cell has attributes 
							{
							//-get the **RowSpanning** if there is a value
							tableAttrValue = String.Empty;
							columnWidth = 0;
							tableAttrValue = node.Attributes.Where(a => a.Name == "rowspan").Single().Value.ToString();
							//-if no **rowspan** is found, set the value to 1
							if(tableAttrValue == String.Empty)
								tableRowSpanQty = 1;
							else
								{
								if(!int.TryParse(tableAttrValue, out tableRowSpanQty))
									tableRowSpanQty = 1;
								}

							//-Get the **Column Spanning** if there is a value
							tableAttrValue = String.Empty;
							tableAttrValue = node.Attributes.Where(a => a.Name == "colspan").Single().Value.ToString();
							if(tableAttrValue == String.Empty)
								tableColumnSpanQty = 1;
							else
								{
								if(!int.TryParse(tableAttrValue, out tableColumnSpanQty))
									tableColumnSpanQty = 1;
								}

							//-Get the column **Width** is specified
							tableAttrValue = node.Attributes.Where(a => a.Name == "style" && a.Value.Contains("width")).Single().Value.ToString();
							if(tableAttrValue == String.Empty)
								columnWidth = 0;
							else
								{
								columnWidth = 0;
								if(tableAttrValue.Contains("%"))
									{
									if(int.TryParse(tableAttrValue.Substring(
										tableAttrValue.IndexOf(":") + 2,
										(tableAttrValue.IndexOf("%") - tableAttrValue.IndexOf(":") - 2)),
										out columnWidth) == false)
										{
										//- Could not parse the integer which means the tableWidth remains as it was before the parse.
										}
									}
								else //-Column width is **NOT** a percentage :. px value
									{
									if(int.TryParse(tableAttrValue.Substring(
										tableAttrValue.IndexOf(":") + 2,
										(tableAttrValue.IndexOf("px") - tableAttrValue.IndexOf(":") - 2)),
										out tableWidth) == false)
										{
										//-Could not parse the integer which means the tableWidth remains as it was before the parse.
										}
									}
								}
							}
						else //- the table Header Cell doesn't have any attributes
							{
							tableColumnSpanQty = 1;
							tableRowSpanQty = 1;
							columnWidth = 0;
							}
						Console.Write("\n\t\t<col> Header Column: Width: {0}\t RowSpan: {1}\t ColumnSpan: {2}\t", columnWidth, tableRowSpanQty, tableColumnSpanQty);
						break;
						}
					//+TableCell - **<td>**
					case "td":
						{
						//-If the **table header** was not determined in the class of the table row *<tr>* tag, then the table has a header row if a *<th>* tag is present
						isHeaderRow = false;

						if(node.HasAttributes)  //- The table Header Cell has attributes 
							{
							foreach(HtmlAttribute tableAttr in node.Attributes)
								{
								switch(tableAttr.Name)
									{
									//-get the **RowSpanning** if there is a value
									case "rowspan": //- get the rowspan if there is any
										{
										if(tableAttr.Value == null)
											tableRowSpanQty = 1;
										else
											{
											if(!int.TryParse(tableAttrValue, out tableRowSpanQty))
												tableRowSpanQty = 1;
											}
										break;
										}
									//-get the **ColumnSpanning** if there is a value
									case "colspan":
										{
										if(tableAttr.Value == null)
											tableColumnSpanQty = 1;
										else
											{
											if(!int.TryParse(tableAttrValue, out tableColumnSpanQty))
												tableColumnSpanQty = 1;
											}
										break;
										}
									//get the column width if it is specified
									case "style":
										{//-Get the column **Width** if it is specified
										if(tableAttr.Value.Contains("width"))
											{
											if(tableAttrValue.Contains("%"))
												{
												if(int.TryParse(tableAttrValue.Substring(
													tableAttrValue.IndexOf(":") + 2,
													(tableAttrValue.IndexOf("%") - tableAttrValue.IndexOf(":") - 2)),
													out columnWidth) == false)
													{
													//- Could not parse the integer which means the tableWidth remains as it was before the parse.
													}
												}
											else //-Column width is **NOT** a percentage :. px value
												{
												if(int.TryParse(tableAttrValue.Substring(
													tableAttrValue.IndexOf(":") + 2,
													(tableAttrValue.IndexOf("px") - tableAttrValue.IndexOf(":") - 2)),
													out tableWidth) == false)
													{
													//-Could not parse the integer which means the tableWidth remains as it was before the parse.
													}
												}
											}
										else
											{
											columnWidth = 0;
											}
										break;
										}
									}
								}
							}
						else //- the table Header Cell doesn't have any attributes
							{
							tableColumnSpanQty = 1;
							tableRowSpanQty = 1;
							columnWidth = 0;
							}
						Console.Write("\n\t\t<col> Header Column: Width: {0}\t RowSpan: {1}\t ColumnSpan: {2}\t", columnWidth, tableRowSpanQty, tableColumnSpanQty);
						break;
						}

					default:
						{
						continue;
						}
					}
				}

			Console.WriteLine("");
			Console.WriteLine("____________________________HTML Ends______________________________________");
			}


		private static string Tabs (int n)
			{
			return new String('\t', n);
	
			}

		private static string CleanText (string parText, string parClientName)
			{
			//!The sequence in which these statements appear is important

			//-keep this code for debugging purposes if strange/unwanted characters appear.
			/*for(int i = 0; i < parText.Length; i++)
				{
				Console.Write("|" + ((int)parText[i]).ToString());
				}
			*/

			string cleanText = parText;
			cleanText = cleanText.Replace(((char)8203).ToString(),"");
			cleanText = cleanText.Replace("&#160;", " "); //-remove *Hard space* characters
			cleanText = cleanText.Replace("     ", " ");		//-cleanup any *5* spaces
			cleanText = cleanText.Replace("   ", " ");		//-cleanup any *triple* spaces
			cleanText = cleanText.Replace("  ", " ");         //-cleanup any *double* spaces
			//cleanText = cleanText.Replace("\r ", "");		//- remove carraige *return* characters
			cleanText = cleanText.Replace("\r", "");		//- remove carraige *return* characters
			//cleanText = cleanText.Replace("\n ", "");		//-remove *New Line* characters
			cleanText = cleanText.Replace("\n", "");          //-remove *New Line* characters
			cleanText = cleanText.Replace("\t", "");          //-remove *Tab* characters
			if(cleanText == " " || cleanText == "  " || cleanText == "   ")
				cleanText = "";                              //-cleanup the string if it contains only a space.

			//-Replace ClientName #tag with actual value
			cleanText = cleanText.Replace("#ClientName#", parClientName);
			cleanText = cleanText.Replace("#clientname#", parClientName);
			cleanText = cleanText.Replace("#CLIENTNAME#", parClientName);

			return cleanText;
			}

		private static Tuple<int, int> GetBulletNumberLevels (string parXpath)
			{
			int bulletLevels = 0;
			int numberLevels = 0;	
			int positionInString = 0;

			//+Check the number of **Headings** </h> tags
			//-Check if the xPath contains any bullets
			positionInString = 0;

			//+Check the number of **Unorganised List** </ul> tags
			//-Check if the xPath contains any bullets
			positionInString = 0;
			if(parXpath.Contains("/ul"))
				{ //- if it contains bullets, count the number of bullets
				for(int i = 0; i < parXpath.Length - 1;)
					{
					//-get the ocurrences of bullets
					positionInString = parXpath.IndexOf("/ul", i);
					if(positionInString >= 0)
						{
						bulletLevels += 1;
						i = positionInString + 3;
						}
					else if(positionInString < 0)
						break;
					}
				}
			else //-If it doesn't contain any bullets, set **bulletLevels** to 0
				{
				bulletLevels = 0;
				}

			//+Check the number of **Organised List** </ol> tags
			//-Check if the xPath contains any **Organised Lists** (Numbered tags).
			positionInString = 0;
			if(parXpath.Contains("/ol"))
				{ //- if it contains any, count the number of occurrences
				for(int i = 0; i < parXpath.Length - 1;)
					{
					//-get the ocurrences of tags
					positionInString = parXpath.IndexOf("/ol", i);
					if(positionInString >= 0)
						{
						numberLevels += 1;
						i = positionInString + 3;
						}
					else if(positionInString < 0)
						break;

					}
				}
			else //-If it doesn't contain any numbers, set **numberLevels** to 0
				{
				numberLevels = 0;
				}

			//-Return the counted occurrences
			return new Tuple<int, int>(bulletLevels, numberLevels);
			}
		}
	}
