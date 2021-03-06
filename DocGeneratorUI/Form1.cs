﻿using System;
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
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Threading;
using System.Windows.Forms;
using DocGeneratorCore;

namespace DocGeneratorUI
	{
	public partial class Form1:Form
		{
		#region Variables
		private static readonly Object lockThreadDatRefresh = new Object();
		DocGeneratorCore.SDDPServiceReference.DesignAndDeliveryPortfolioDataContext sddpDataContext;
		DocGeneratorCore.enumPlatform platform;

		#endregion

		#region Properties
		public enumDocumentTypes DocumentType { get; private set; }
		public Boolean UnhandledError { get; private set; }
		public List<string> ErrorMessages { get; private set; } = new List<string>();
		public enumDocumentStatusses DocumentStatus { get; private set; }
		public bool HyperlinkEdit { get; private set; }
		public Boolean HyperlinkView { get; private set; }
		public String Template { get; private set; }
		public String LocalDocumentURI { get; private set; }
		public String FileName { get; private set; }
		public uint PageWith { get; set; }
		public uint PageHeight { get;  set; }
		public Boolean ColorCodingLayer1 { get; private set; }
		public Boolean ColorCodingLayer2 { get; private set; }
		public Boolean Introductory_Section { get; private set; }
		public Boolean Introduction { get; private set; }

		#endregion

		#region Methods
		public Form1()
			{
			Thread.CurrentThread.Name = "MainGUI";
			InitializeComponent();
			}

		private void Form1_Load(object sender, EventArgs e)
			{

			if(Properties.Settings.Default.LastSelectedPlatform.IndexOf(value: "prod", startIndex: 0, comparisonType: StringComparison.OrdinalIgnoreCase) >= 0)
				{
				this.radioButtonPROD.Checked = true;
				}
			else if (Properties.Settings.Default.LastSelectedPlatform.IndexOf(value: "q", startIndex: 0, comparisonType: StringComparison.OrdinalIgnoreCase) >= 0)
				{
				this.radioButtonQA.Checked = true;
				}
			else if (Properties.Settings.Default.LastSelectedPlatform.IndexOf(value: "dev", startIndex: 0, comparisonType: StringComparison.OrdinalIgnoreCase) >= 0)
				{
				this.radioButtonDEV.Checked = true;
				}
			
			Application.DoEvents();
			}

		private void buttonGenerateAll_Click(object sender, EventArgs e)
			{

			if (this.radioButtonDEV.Checked)
				platform = enumPlatform.Development;
			else if (this.radioButtonPROD.Checked)
				platform = enumPlatform.Production;
			else if (this.radioButtonQA.Checked)
				platform = enumPlatform.QualityAssurance;
			else
				{
				MessageBox.Show(
					text: "Please specify a Platform before proceeding.",
					caption: "No Platform specified",
					buttons: MessageBoxButtons.OK,
					icon: MessageBoxIcon.Error);
				this.groupBoxPlatform.Focus();
				return;
				}

			Console.WriteLine("Platform: {0}", platform);
			Cursor.Current = Cursors.WaitCursor;

			//-| Initialise the listDocumentCollections object if it is null.
			List<DocumentCollection> listDocumentCollections = new List<DocumentCollection>();
			string strExceptionMessage = string.Empty;
			try
				{
				var dsDocCollections = 
					from dsDocumentCollection in sddpDataContext.DocumentCollectionLibrary
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
					//-| Add the DocumentCollection object to the listDocumentCollection
					listDocumentCollections.Add(objDocumentCollection);
					}

				//-| Check if there are any Document Collections to generate
				if(listDocumentCollections.Count > 0)
					{
					foreach(DocumentCollection entryDocumentCollection in listDocumentCollections)
						{
						//-| Invoke the DocGeneratorCore's MainController object MainProcess method
						MainController objMainController = new MainController();
						objMainController.DocumentCollectionsToGenerate = listDocumentCollections;
						objMainController.Platform = Properties.Settings.Default.CurrentPlatform;
						objMainController.MainProcess();
						}
					}

				MessageBox.Show("Successfully completed a cycle run of the DocGeneratorCore controller and processed all unprocessed "
				+ " Document Collections" + " \nClick it again to run another cycle.",
				"Run successfully completed.", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
			catch(DataServiceClientException exc)
				{
				strExceptionMessage = "*** Exception ERROR ***: DocGeneratorServer cannot access site: " 
					+ Properties.Settings.Default.CurrentURLSharePoint 
					+ Properties.Settings.Default.CurrentURLSharePointSitePortion
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
					+ Properties.Settings.Default.CurrentURLSharePoint
					+ Properties.Settings.Default.CurrentURLSharePointSitePortion
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
					+ Properties.Settings.Default.CurrentURLSharePoint
					+ Properties.Settings.Default.CurrentURLSharePointSitePortion
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
					+ Properties.Settings.Default.CurrentURLSharePoint
					+ Properties.Settings.Default.CurrentURLSharePointSitePortion
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
					+ Properties.Settings.Default.CurrentURLSharePoint
					+ Properties.Settings.Default.CurrentURLSharePointSitePortion
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
						+ Properties.Settings.Default.CurrentURLSharePoint
						+ Properties.Settings.Default.CurrentURLSharePointSitePortion
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
						+ Properties.Settings.Default.CurrentURLSharePoint
						+ Properties.Settings.Default.CurrentURLSharePointSitePortion
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
		private void buttonGenerateDocCollection_Click(object sender, EventArgs e)
			{
			//-| Validate the input
			if (this.radioButtonDEV.Checked)
				platform = enumPlatform.Development;
			else if (this.radioButtonPROD.Checked)
				platform = enumPlatform.Production;
			else if (this.radioButtonQA.Checked)
				platform = enumPlatform.QualityAssurance;
			else
				{
				MessageBox.Show(
					text: "Please specify a Platform before proceeding.",
					caption: "No Platform specified",
					buttons: MessageBoxButtons.OK,
					icon: MessageBoxIcon.Error);
				this.groupBoxPlatform.Focus();
				return;
				}

			Console.WriteLine("Platform: {0}", platform);

			if (string.IsNullOrEmpty(this.textBoxDocumentCollection.Text))
				{
				MessageBox.Show("Please enter a numeric value, before clicking the Generate button",
						"No value entred.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
				}
			else
				Console.WriteLine("Generate documents for DocumentCollection: {0}", this.textBoxDocumentCollection.Text);

			int intDocumentCollectionID = 0;

			bool bIsNumeric = int.TryParse(this.textBoxDocumentCollection.Text.Trim(), out intDocumentCollectionID);

			if(!bIsNumeric)
				{
				MessageBox.Show("Only numeric values can be entered.",
						"Invalid value entered.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
				}
			//-|Begin to process...
			Cursor.Current = Cursors.WaitCursor;
		
			Application.DoEvents();

			string strExceptionMessage = String.Empty;
			//-|Initialise the DocumentCollections List
			List<DocumentCollection> listDocumentCollections = new List<DocumentCollection>();

			try
				{
				//-|Read the Document Collection from SharePoint and load it into the Object model.
				var dsDocumentCollections = 
					from dsDC in sddpDataContext.DocumentCollectionLibrary
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
					//-|Create a DocumentCollection instance and populate the basic attributes.
					Application.DoEvents();
					DocumentCollection objDocumentCollection = new DocumentCollection();
					objDocumentCollection.ID = objDocCollection.Id;
					if(objDocCollection.Title == null)
						objDocumentCollection.Title = "Collection Title for entry " + objDocCollection.Id;
					else
						objDocumentCollection.Title = objDocCollection.Title;

					objDocumentCollection.DetailComplete = false;
					//-|Add the DocumentCollection object to the DocumentCollection List
					listDocumentCollections.Add(objDocumentCollection);
					}
				Application.DoEvents();
				
				//-|Check if there are any Document Collections to generate
				if(listDocumentCollections.Count > 0)
					{
					//-| Invoke the DocGeneratorCore's **MainController** object and send all the Document Collections entries for generation.
					MainController objMainController = new MainController();
					objMainController.DocumentCollectionsToGenerate = listDocumentCollections;
					objMainController.Platform = Properties.Settings.Default.CurrentPlatform;
					objMainController.MainProcess();
					}
				}
			catch(DataServiceClientException exc)
				{
				strExceptionMessage = "*** Exception ERROR ***: DocGeneratorUI cannot access site: " 
					+ Properties.Settings.Default.CurrentURLSharePoint
					+ Properties.Settings.Default.CurrentURLSharePointSitePortion
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
						+ Properties.Settings.Default.CurrentURLSharePoint
						+ Properties.Settings.Default.CurrentURLSharePointSitePortion
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
					+ Properties.Settings.Default.CurrentURLSharePoint
					+ Properties.Settings.Default.CurrentURLSharePointSitePortion
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
					+ Properties.Settings.Default.CurrentURLSharePoint
					+ Properties.Settings.Default.CurrentURLSharePointSitePortion
					+ " Please check that the computer/server is connected to the Domain network "
					+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
					+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
				Application.DoEvents();
				Cursor.Current = Cursors.Default;
				MessageBox.Show(strExceptionMessage,
					"Unable to generatate any documents.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
				}
			catch (LocalDatabaseCreationExeption exc)
				{
				strExceptionMessage = "*** Exception ERROR ***: Cannot create the local Database for " + platform + "\n " + exc.Message;
				Application.DoEvents();
				Cursor.Current = Cursors.Default;
				MessageBox.Show(strExceptionMessage,
					"Unable to generatate any documents.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
				}

			catch (Exception exc)
				{
				if(exc.HResult == -2146330330)
					{
					strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " 
						+ Properties.Settings.Default.CurrentURLSharePoint
						+ Properties.Settings.Default.CurrentURLSharePointSitePortion
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
						+ Properties.Settings.Default.CurrentURLSharePoint
						+ Properties.Settings.Default.CurrentURLSharePointSitePortion
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
						+ Properties.Settings.Default.CurrentURLSharePoint
						+ Properties.Settings.Default.CurrentURLSharePointSitePortion
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

			MessageBox.Show("Successfully completed the generation of Document Collection: " + this.textBoxDocumentCollection.Text
				+ " \nClick it again to generate the same Document Collection or enter another number...",
				"Generation successfully completed.", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}


		private void buttonSendEmail_Click(object sender, EventArgs e)
			{
			if (this.radioButtonDEV.Checked
			|| this.radioButtonPROD.Checked
			|| this.radioButtonQA.Checked)
				{ }
			else
				{
				MessageBox.Show(
					text: "Please specify a Platform before proceeding.",
					caption: "No Platform specified",
					buttons: MessageBoxButtons.OK,
					icon: MessageBoxIcon.Error);
				this.groupBoxPlatform.Focus();
				return;
				}


			// Validate the e-mail input
			if (textBoxEmail.Text == null || textBoxEmail.Text == "")
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
					parReceipient: this.textBoxEmail.Text,
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
			if (this.radioButtonDEV.Checked
			|| this.radioButtonPROD.Checked
			|| this.radioButtonQA.Checked)
				{ }
			else
				{
				MessageBox.Show(
					text: "Please specify a Platform before proceeding.",
					caption: "No Platform specified",
					buttons: MessageBoxButtons.OK,
					icon: MessageBoxIcon.Error);
				this.groupBoxPlatform.Focus();
				return;
				}

			// Validate the e-mail input
			if (textBoxEmail.Text == null || textBoxEmail.Text == "")
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
					parReceipient: this.textBoxEmail.Text,
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
			if (this.radioButtonDEV.Checked
			|| this.radioButtonPROD.Checked
			|| this.radioButtonQA.Checked)
				{ }
			else
				{
				MessageBox.Show(
					text: "Please specify a Platform before proceeding.",
					caption: "No Platform specified",
					buttons: MessageBoxButtons.OK,
					icon: MessageBoxIcon.Error);
				this.groupBoxPlatform.Focus();
				return;
				}

			// Validate the e-mail input
			if (textBoxEmail.Text == null || textBoxEmail.Text == "")
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
					parReceipient: this.textBoxEmail.Text,
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

		private void buttonTestHTML_Click(Object sender, EventArgs e)
			{
			//-Declare the htmlData object instance
			HtmlAgilityPack.HtmlDocument htmlData = new HtmlAgilityPack.HtmlDocument();
			//-Load the HTML content into the htmlData object.
			//htmlData.Load(@"E:\Development\Projects\DocGeneratorGUI\DocGeneratorUI\TestData\HTMLtestPage_CascadedBulletList.html");
			//htmlData.Load(@"E:\Development\Projects\DocGeneratorGUI\DocGeneratorUI\TestData\HTMLtestPage_CascadedNumberList.html");
			//htmlData.Load(@"E:\Development\Projects\DocGeneratorGUI\DocGeneratorUI\TestData\HTMLtestPage_Everything.html");
			htmlData.Load(@"E:\Development\Projects\DocGeneratorGUI\DocGeneratorUI\TestData\HTMLtestPage_ProblemContent.html");
			//-Load a string into the HTMLDocument
			Console.WriteLine("______________________________HTML starts______________________________________");
			//htmlData.LoadHtml(htmlText);
			Console.WriteLine("HTML: {0}", htmlData.DocumentNode.OuterHtml);
			Console.WriteLine("_______________________________________________________________________________");
			//- Set the ROOT of the data loaded in the htmlData
			var htmlRoot = htmlData.DocumentNode;
			Console.WriteLine("Root Node Tag..............: {0}", htmlRoot.Name);
			this.ErrorMessages = new List<string>();
			/*Console.WriteLine("\t - Has Child Nodes......? {0}", htmlRoot.HasChildNodes);
			*/

			//- The sequence of the values in the tuple variable is: 
			//-**1** *Bullet Levels*, 
			//-**2** *Number Levels* 
			Tuple<int, int> bulletNumberLevels = new Tuple<int, int>(0, 0);
			int bulletLevel = 0;
			int numberLevel = 0;
			int headingLevel = 0;
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

				//Console.WriteLine(">-- {0} --<", node.Name);

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
							cleanText = HTMLdecoder.CleanText(node.InnerText, clientName);
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
							cleanText = HTMLdecoder.CleanText(node.InnerText, clientName);
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
						//-Get the number of bullet - and number - levels in the xPath
						bulletNumberLevels = HTMLdecoder.GetBulletNumberLevels(node.XPath);
						bulletLevel = bulletNumberLevels.Item1;
						numberLevel = bulletNumberLevels.Item2;

						//-Check if the paragraph is part of a bullet- number- list
						if(numberLevel > 0 || bulletLevel > 0)
							{
							//-don't insert a paragraph, because the text is part of the bullet or number
							break;
							}

						if(node.HasChildNodes)
							{
							Console.Write("\n\t <{0}> ", node.Name);
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
						bulletNumberLevels = HTMLdecoder.GetBulletNumberLevels(node.XPath);
						bulletLevel = bulletNumberLevels.Item1;
						numberLevel = bulletNumberLevels.Item2;
						
						if(node.HasChildNodes)
							{
							if(bulletLevel > 0)
								{
								Console.Write("\n {0} - <{1}>", Tabs(headingLevel + bulletLevel), node.Name);
								}
							else if(numberLevel > 0)
								{
								Console.Write("\n {0} {1}. <{2}>", Tabs(headingLevel  + numberLevel), numberLevel, node.Name);
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
						//Console.WriteLine("\t ~~~ skip {0} ~~~", node.Name);
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



		private void buttonGenerateTestDoc_Click(Object sender, EventArgs e)
			{

			this.DocumentType = enumDocumentTypes.ISD_Document_DRM_Inline;
			Console.WriteLine("\t Begin to generate {0}", this.DocumentType);
			string parClientName = "XYZ Company";
			this.ErrorMessages = new List<string>();
			this.UnhandledError = false;
			DateTime timeStarted = DateTime.Now;
			string hyperlinkImageRelationshipID = "";
			string documentCollection_HyperlinkURL = "";
			string currentListURI = "";
			string currentHyperlinkViewEditURI = "";
			string currentContentLayer = "None";
			bool layerHeadingWritten = false;
			bool drmHeading = false;
			Table objActivityTable = new Table();
			Table objServiceLevelTable = new Table();
			int? intLayer1upElementID = 0;
			int? intLayer1upDeliverableID = 0;
			int intTableCaptionCounter = 0;
			int intImageCaptionCounter = 0;
			int numberingCounter = 49;
			int iPictureNo = 49;
			int intHyperlinkCounter = 9;
			string htmlString = String.Empty;

			this.HyperlinkEdit = false;
			this.Template = "InternalServiceDefinitionTemplate.dotx";
			this.Introductory_Section = true;
			this.Introduction = true;
			try
				{
				if(this.HyperlinkEdit)
					{
					documentCollection_HyperlinkURL = "https://en.wikipedia.org/wiki/Pixel_density";
					currentHyperlinkViewEditURI = documentCollection_HyperlinkURL;
					}
				if(this.HyperlinkView)
					{
					documentCollection_HyperlinkURL = "https://en.wikipedia.org/wiki/Microsoft_Office_XML_formats";
					currentHyperlinkViewEditURI = documentCollection_HyperlinkURL;
					}

				// define a new objOpenXMLdocument
				oxmlDocument objOXMLdocument = new oxmlDocument();
				// use CreateDocumentFromTemplate method to create a new MS Word Document based on
				// the relevant template
				if(objOXMLdocument.CreateDocWbkFromTemplate(
					parDocumentOrWorkbook: enumDocumentOrWorkbook.Document,
					parTemplateURL: this.Template,
					parDocumentType: this.DocumentType,
					parSDDPdataContext: sddpDataContext))
					{
					Console.WriteLine("\t\t objOXMLdocument:\n" +
					"\t\t\t+ LocalDocumentPath: {0}\n" +
					"\t\t\t+ DocumentFileName.: {1}\n" +
					"\t\t\t+ DocumentURI......: {2}", objOXMLdocument.LocalPath, objOXMLdocument.Filename, objOXMLdocument.LocalURI);
					}
				else
					{
					//- if the file creation failed.
					throw new DocumentUploadException(message: "DocGenerator was unable to create the document based on the template.");
					}

				this.LocalDocumentURI = objOXMLdocument.LocalURI;
				this.FileName = objOXMLdocument.Filename;

				this.DocumentStatus = enumDocumentStatusses.Creating;
				// Open the MS Word document in Edit mode
				WordprocessingDocument objWPdocument = WordprocessingDocument.Open(path: objOXMLdocument.LocalURI, isEditable: true);
				// Define all open XML object to use for building the document
				MainDocumentPart objMainDocumentPart = objWPdocument.MainDocumentPart;
				Body objBody = objWPdocument.MainDocumentPart.Document.Body;          // Define the objBody of the document
				Paragraph objParagraph = new Paragraph();
				ParagraphProperties objParaProperties = new ParagraphProperties();
				Run objRun = new Run();
				RunProperties objRunProperties = new RunProperties();
				Text objText = new Text();
				// Declare the HTMLdecoder object and assign the document's WordProcessing Body to
				// the WPbody property.
				//original version
				//DocGeneratorCore.HTMLdecoder objHTMLdecoder = new DocGeneratorCore.HTMLdecoder();
				//new version
				HTMLdecoder objHTMLdecoder = new HTMLdecoder();
				//!Set these HTMLdecoder properties....
				//-Set the properties of the WordProcessing Body object (WPbody) of the HTMLdecoder object instance.
				objHTMLdecoder.WPbody = objBody;
				objHTMLdecoder.ClientName = parClientName;

				// Determine the Page Size for the current Body object.
				SectionProperties objSectionProperties = new SectionProperties();
				this.PageWith = 0; // Convert.ToUInt32(Properties.AppResources.DefaultPageWidth);
				this.PageHeight = 0; //Convert.ToUInt32(Properties.AppResources.DefaultPageHeight);

				if (objBody.GetFirstChild<SectionProperties>() != null)
					{
					objSectionProperties = objBody.GetFirstChild<SectionProperties>();
					PageSize objPageSize = objSectionProperties.GetFirstChild<PageSize>();
					PageMargin objPageMargin = objSectionProperties.GetFirstChild<PageMargin>();
					if (objPageSize != null)
						{
						this.PageWith = objPageSize.Width;
						this.PageHeight = objPageSize.Height;
						//Console.WriteLine("\t\t Page width x height: {0} x {1} twips", this.PageWith, this.PageHight);
						}
					if (objPageMargin != null)
						{
						if (objPageMargin.Left != null)
							{
							this.PageWith -= objPageMargin.Left;
							//Console.WriteLine("\t\t\t - Left Margin..: {0} twips", objPageMargin.Left);
							}
						if (objPageMargin.Right != null)
							{
							this.PageWith -= objPageMargin.Right;
							//Console.WriteLine("\t\t\t - Right Margin.: {0} twips", objPageMargin.Right);
							}
						if (objPageMargin.Top != null)
							{
							string tempTop = objPageMargin.Top.ToString();
							//Console.WriteLine("\t\t\t - Top Margin...: {0} twips", tempTop);
							this.PageHeight -= Convert.ToUInt32(tempTop);
							}
						if (objPageMargin.Bottom != null)
							{
							string tempBottom = objPageMargin.Bottom.ToString();
							//Console.WriteLine("\t\t\t - Bottom Margin: {0} twips", tempBottom);
							this.PageHeight -= Convert.ToUInt32(tempBottom);
							}
						}
					}
				// Subtract the Table/Image Left indentation value from the Page width to ensure the
				// table/image fits in the available space.
				this.PageWith -= 855; //Convert.ToUInt16(Properties.AppResources.Document_Table_Left_Indent);
				

				//!Set the HTMLdecoder's **PageHeight** and **PageWidth** properties
				objHTMLdecoder.PageHeightDxa = this.PageHeight;
				objHTMLdecoder.PageWidthDxa = this.PageWith;

				// Check whether Hyperlinks need to be included and add the image to the Document Body
				if(this.HyperlinkEdit || this.HyperlinkView)
					{
					//Insert and embed the hyperlink image in the document and keep the Image's Relationship ID in a variable for repeated use
					hyperlinkImageRelationshipID = oxmlDocument.Insert_HyperlinkImage(parMainDocumentPart: ref objMainDocumentPart,
					parSDDPdatacontext: sddpDataContext);
					}

				//Check is Content Layering was requested and add a Ledgend for the colour coding of content
				if(this.ColorCodingLayer1 || this.ColorCodingLayer2)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 0, parNoNumberedHeading: true);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: "Colour coding",
						parBold: true);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 0);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: "The following colour coding will appear in the document and it has the following meaning:");
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					objParagraph = oxmlDocument.Construct_BulletParagraph(parBulletLevel: 1);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: "Services Framework Text is this colour.",
						parContentLayer: "Layer1");
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					objParagraph = oxmlDocument.Construct_BulletParagraph(parBulletLevel: 1);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: "Services Framework Plus is this colour.",
						parContentLayer: "Layer2");
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 0);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: " ");
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);
					}

				this.DocumentStatus = enumDocumentStatusses.Building;

				//++ Insert a Section
				if(this.Introductory_Section)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: "DD Heading 1",
						parIsNewSection: true);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);
					}

				//+ Insert a Heading
				if(this.Introduction)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
					objRun = oxmlDocument.Construct_RunText(parText2Write: "DD Heading 2");
					// Check if a hyperlink must be inserted
					if(this.HyperlinkEdit || this.HyperlinkView)
						{
						intHyperlinkCounter += 1;
						Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
							parMainDocumentPart: ref objMainDocumentPart,
							parImageRelationshipId: hyperlinkImageRelationshipID,
							parClickLinkURL: documentCollection_HyperlinkURL,
							parHyperlinkID: intHyperlinkCounter);
						objRun.Append(objDrawing);
						}
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					//+Load the HTML data from a file in the TestData directory
					string myPath = @"E:\Development\Projects\DocGeneratorGUI\DocGeneratorUI\TestData\";
					//string htmlFilename = @myPath + "HTMLtestPage_ComplexTable.html";
					//string htmlFilename = @myPath + "HTMLtestPage_Everything.html";
					//string htmlFilename = @myPath + "HTMLtestPage_ParagraphNumbering.html";
					string htmlFilename = @myPath + "HTMLtestPage_ProblemContent.html";
					//string htmlFilename = @myPath + "HTMLtestPage_ProblemContent_All.html";

					htmlString = System.IO.File.ReadAllText(@htmlFilename);

					if(htmlString != null)
						{
						try
							{
														
							objHTMLdecoder.DecodeHTML(
								parMainDocumentPart: ref objMainDocumentPart,
								parDocumentLevel: 1,
								parTableCaptionCounter: ref intTableCaptionCounter,
								parImageCaptionCounter: ref intImageCaptionCounter,
								parNumberingCounter: ref numberingCounter,
								parPictureNo: ref iPictureNo,
								parClientName: parClientName,
								parPageHeightDxa: this.PageHeight,
								parPageWidthDxa: this.PageWith,
								parHyperlinkID: ref intHyperlinkCounter,
								parSharePointSiteURL: @"https:\\teams.dimensiondata.com\sites\servicecatalogue\",
								parHTML2Decode: HTMLdecoder.CleanHTML(htmlString, parClientName));
							}
						catch(InvalidContentFormatException exc)
							{
							Console.WriteLine("\nException occurred: {0}", exc.Message);
							// A Table content error occurred, record it in the error log.
							this.ErrorMessages.Add("Content Error: The Document Collection ID: " + 1
								+ " contains a CONTENT error in Introduction's Enhance Rich Text. " + exc.Message);
							objParagraph = oxmlDocument.Construct_Error(parText: "A content error occurred at this position and valid content could " +
								"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it. Error Detail: "
								+ exc.Message);
							if(this.HyperlinkEdit || this.HyperlinkView)
								{
								intHyperlinkCounter += 1;
								Drawing objDrawing = oxmlDocument.Construct_ClickLinkHyperlink(
									parMainDocumentPart: ref objMainDocumentPart,
									parImageRelationshipId: hyperlinkImageRelationshipID,
									parHyperlinkID: intHyperlinkCounter,
									parClickLinkURL: documentCollection_HyperlinkURL);
								objRun.Append(objDrawing);
								}
							objBody.Append(objParagraph);
							}
						}
					}

				//++ Insert the Document Generation Error Section
				if(this.ErrorMessages.Count > 0)
					{
					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 1);
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: "Errors that occurred in the HTML",
						parIsNewSection: true);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					objParagraph = oxmlDocument.Construct_Heading(parHeadingLevel: 2);
					objRun = oxmlDocument.Construct_RunText(parText2Write: "Errors that occurred");
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					foreach(var errorMessageEntry in this.ErrorMessages)
						{
						objParagraph = oxmlDocument.Construct_Error(parText: errorMessageEntry);
						objBody.Append(objParagraph);
						}
					}

				//+ Validate the document with OpenXML validator
				OpenXmlValidator objOXMLvalidator = new OpenXmlValidator(fileFormat: DocumentFormat.OpenXml.FileFormatVersions.Office2010);
				int errorCount = 0;
				Console.WriteLine("\n\rValidating document....");
				foreach(ValidationErrorInfo validationError in objOXMLvalidator.Validate(objWPdocument))
					{
					errorCount += 1;
					Console.WriteLine("------------- # {0} -------------", errorCount);
					Console.WriteLine("Error ID................: {0}", validationError.Id);
					Console.WriteLine("Description.............: {0}", validationError.Description);
					Console.WriteLine("Error Type..............: {0}", validationError.ErrorType);
					Console.WriteLine("Error Part..............: {0}", validationError.Part.Uri);
					Console.WriteLine("Root Element............: {0}", validationError.Part.RootElement);
					Console.WriteLine("Error Path..............: {0}", validationError.Path.XPath);
					Console.WriteLine("Error Path PartUri......: {0}", validationError.Path.PartUri);
					Console.WriteLine("Error Node..............: {0}", validationError.Node);
					Console.WriteLine("Error Related Node......: {0}", validationError.RelatedNode);
					Console.WriteLine("Node Local Name.........: {0}", validationError.Node.LocalName);
					Console.WriteLine("Node Prefix.............: {0}", validationError.Node.Prefix);
					Console.WriteLine("Node XML Qualifier Name.: {0}", validationError.Node.XmlQualifiedName);
					}

				Console.WriteLine("Document generation completed, saving and closing the document.");
				// Save and close the Document
				objWPdocument.Close();

				this.DocumentStatus = enumDocumentStatusses.Completed;

				Console.WriteLine(
					"Generation started...: {0} \nGeneration completed: {1} \n Durarion..........: {2}",
					timeStarted, DateTime.Now, (DateTime.Now - timeStarted));

				} // end Try

			//++ -------------------
			//++ Handle Exceptions
			//++ -------------------
			//+ NoContentspecified Exception
			catch(NoContentSpecifiedException exc)
				{
				this.ErrorMessages.Add(exc.Message);
				this.DocumentStatus = enumDocumentStatusses.Error;
				return; //- exit the method because there is no files to cleanup
				}

			//+ UnableToCreateDocument Exception
			catch(UnableToCreateDocumentException exc)
				{
				this.ErrorMessages.Add(exc.Message);
				this.DocumentStatus = enumDocumentStatusses.FatalError;
				return; //- exit the method because there is no files to cleanup
				}

			//+ DocumentUpload Exception
			catch(DocumentUploadException exc)
				{
				this.ErrorMessages.Add(exc.Message);
				this.DocumentStatus = enumDocumentStatusses.FatalError;
				}

			//+ OpenXMLPackage Exception
			catch(OpenXmlPackageException exc)
				{
				this.ErrorMessages.Add("Unfortunately, an unexpected error occurred during document generation and the document could not be produced. ["
					+ "[OpenXMLPackageException: " + exc.HResult + "Detail: " + exc.Message + "]");
				this.DocumentStatus = enumDocumentStatusses.FatalError;
				this.UnhandledError = true;
				}

			//+ ArgumentNull Exception
			catch(ArgumentNullException exc)
				{
				this.ErrorMessages.Add("Unfortunately, an unexpected error occurred during  ocument generation and the document could not be produced. ["
					+ "[ArgumentNullException: " + exc.HResult + "Detail: " + exc.Message + "]");
				this.DocumentStatus = enumDocumentStatusses.FatalError;
				this.UnhandledError = true;
				}

			//+ Exception (any not specified Exception)
			catch(Exception exc)
				{
				this.ErrorMessages.Add("An unexpected error occurred during the document generation and the document could not be produced. ["
					+ "[Exception: " + exc.HResult + "Detail: " + exc.Message + "]");
				this.DocumentStatus = enumDocumentStatusses.FatalError;
				this.UnhandledError = true;
				}

			Console.WriteLine("{1} End of the generation of {0} {1}", this.DocumentType, new string('#', 25));

			}

		private void buttonTestSomething_Click(object sender, EventArgs e)
			{
			int intValue = 0;
			Console.Write("\n\nExtract the Width from ");
			string testString = "margin&#58;5px;width&#58;50%;height&#58;14%;";
			testString = "width&#58;80px;";
			string myString = testString;
			Console.WriteLine(myString);
			string WorkString = myString.Replace("&#58;", "");
			WorkString = WorkString.Replace(":", "");
			WorkString = WorkString.Replace(" ", "");
			int startInString = WorkString.IndexOf(value: "width", startIndex: 0) + 5;
			int endInString = WorkString.IndexOf(";", startIndex: WorkString.IndexOf("width", 0) + 5);

			Console.WriteLine("Begin Posisie:{0} = {1}", startInString, WorkString.Substring(startInString,1));
			Console.WriteLine("End Posisie..:{0} = {1}", endInString, WorkString.Substring(endInString, 1));
			Console.WriteLine("{0} - {1}", endInString, startInString);
			Console.WriteLine("{0}", WorkString.Substring(startInString, 
				length: endInString - startInString ));
			if (WorkString.Substring(
				startIndex: startInString,
				length: endInString - startInString).Contains("px") == true)
				{
				Console.Write("\t = Pixels: ");
				WorkString = WorkString.Substring(startIndex: startInString, length: endInString - startInString -2);
				Console.Write(WorkString);
				int.TryParse(WorkString, out intValue);
				}
			else
				{
				Console.Write("\t = Percentage: ");
				WorkString = WorkString.Substring(startIndex: startInString, length: endInString - startInString - 1);
				Console.Write(WorkString);
				int.TryParse(WorkString, out intValue);
				}


			Console.Write("\n\nExtract the Height from :");
			myString = testString;
			Console.Write(myString);
			WorkString = myString.Replace("&#58;", "");
			WorkString = WorkString.Replace(":", "");
			WorkString = WorkString.Replace(" ", "");
			/*
			Console.WriteLine("Height starts at: {0}", WorkString.IndexOf("height", 0));
			Console.WriteLine("Height VALUE starts at: {0} = {1}", WorkString.IndexOf("height", 0) + 6, WorkString.Substring(WorkString.IndexOf("height", 0) + 6,
				5));
			Console.WriteLine("Height ends at: {0}", WorkString.IndexOf(value: ";", startIndex: WorkString.IndexOf("height", 0) + 6));
			Console.WriteLine("{0} - {1}", WorkString.IndexOf(value: ";", startIndex: WorkString.IndexOf("height", 0) + 6), WorkString.IndexOf("height", 0) + 6);
			Console.WriteLine("Height value: {0}", WorkString.Substring(startIndex: WorkString.IndexOf("height", 0) + 6,
				length: (WorkString.IndexOf(value: ";", startIndex: WorkString.IndexOf("height", 0) + 6)) - (WorkString.IndexOf("height", 0) + 6) ) );
				*/
			if (WorkString.Substring(startIndex: WorkString.IndexOf("height", 0) + 6,
				length: (WorkString.IndexOf(value: ";", startIndex: WorkString.IndexOf("height", 0) + 6)) - (WorkString.IndexOf("height", 0) + 6)).Contains("px"))
				{
				Console.Write("\t = Pixels: ");
				WorkString = WorkString.Substring(
					startIndex: WorkString.IndexOf(value: "height", startIndex: 0) + 6,
					length: (WorkString.IndexOf(";", startIndex: WorkString.IndexOf("height", 0) + 6) - 2) - (WorkString.IndexOf("height", 0) + 6));
				Console.Write(WorkString);
				int.TryParse(WorkString, out intValue);
				}
			else
				{
				Console.Write("\t = Percentage: ");
				WorkString = WorkString.Substring(
					startIndex: WorkString.IndexOf(value: "height", startIndex: 0) + 6,
					length: (WorkString.IndexOf(";", startIndex: WorkString.IndexOf("height", 0) + 6) - 1) - (WorkString.IndexOf("height", 0) + 6));
				Console.Write(WorkString);
				int.TryParse(WorkString, out intValue);
				}
			}

		private void radioButtonPROD_CheckedChanged(object sender, EventArgs e)
			{
			if(this.radioButtonPROD.Checked)
				{
				Properties.Settings.Default.LastSelectedPlatform = "Production";
				Properties.Settings.Default.Save();
				Properties.Settings.Default.CurrentDatabaseHost = Dns.GetHostName();
				Properties.Settings.Default.CurrentPlatform = DocGeneratorCore.enumPlatform.Production.ToString();
				Properties.Settings.Default.CurrentDatabaseLocation = Properties.Settings.Default.DatabaseLocationPROD;
				Properties.Settings.Default.CurrentSDDPwebReference = Properties.Settings.Default.SDDPwebReferencePROD;
				Properties.Settings.Default.CurrentURLSharePoint = Properties.Settings.Default.URLSharePointPROD;
				Properties.Settings.Default.CurrentURLSharePointSitePortion = Properties.Settings.Default.URISharePointSitePortionPROD;
				SetDataContext();
				
				}
			}

		private void radioButtonQA_CheckedChanged(object sender, EventArgs e)
			{
			if (this.radioButtonQA.Checked)
				{
				Properties.Settings.Default.LastSelectedPlatform = "QA";
				Properties.Settings.Default.Save();
				Properties.Settings.Default.CurrentDatabaseHost = Dns.GetHostName();
				Properties.Settings.Default.CurrentPlatform = DocGeneratorCore.enumPlatform.QualityAssurance.ToString();
				Properties.Settings.Default.CurrentDatabaseLocation = Properties.Settings.Default.DatabaseLocationQA;
				Properties.Settings.Default.CurrentSDDPwebReference = Properties.Settings.Default.SDDPwebReferenceQA;
				Properties.Settings.Default.CurrentURLSharePoint = Properties.Settings.Default.URLSharePointQA;
				Properties.Settings.Default.CurrentURLSharePointSitePortion = Properties.Settings.Default.URISharePointSitePortionQA;
				SetDataContext();
				}
			}

		private void radioButtonDEV_CheckedChanged(object sender, EventArgs e)
			{
			if (this.radioButtonDEV.Checked)
				{
				Properties.Settings.Default.LastSelectedPlatform = "Development";
				Properties.Settings.Default.Save();
				Properties.Settings.Default.CurrentPlatform = DocGeneratorCore.enumPlatform.Development.ToString();
				Properties.Settings.Default.CurrentDatabaseLocation = Properties.Settings.Default.DatabaseLocationDEV;
				Properties.Settings.Default.CurrentSDDPwebReference = Properties.Settings.Default.SDDPwebReferenceDEV;
				Properties.Settings.Default.CurrentURLSharePoint = Properties.Settings.Default.URLSharePointDEV;
				Properties.Settings.Default.CurrentURLSharePointSitePortion = Properties.Settings.Default.URISharePointSitePortionDEV;
				SetDataContext();
				}
			}

		private void SetDataContext()
			{
			//- Initiate the SharePoint Datacontext...
			sddpDataContext = new DocGeneratorCore.SDDPServiceReference.DesignAndDeliveryPortfolioDataContext(new
					Uri(Properties.Settings.Default.CurrentURLSharePoint
							+ Properties.Settings.Default.CurrentURLSharePointSitePortion
							+ Properties.Settings.Default.SharePointRESTuri));

			sddpDataContext.Credentials = new NetworkCredential(
				userName: Properties.Resources.DocGenerator_AccountName,
				password: Properties.Resources.DocGenerator_Account_Password,
				domain: Properties.Resources.DocGenerator_AccountDomain);

			sddpDataContext.MergeOption = MergeOption.NoTracking;

			}
		#endregion
		}
	}
