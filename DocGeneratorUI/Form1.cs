using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Services.Client;
using System.Drawing;
using System.Linq;
using System.Net;
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
		}
	}
