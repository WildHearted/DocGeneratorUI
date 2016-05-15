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
			//- Initiate the SharePoint Datacontext...
			this.completeDataSet.SDDPdatacontext = new DocGeneratorCore.SDDPServiceReference.DesignAndDeliveryPortfolioDataContext(
					new Uri(Properties.Resources.SharePointSiteURL + Properties.Resources.SharePointRESTuri));

			this.completeDataSet.SDDPdatacontext.Credentials = new NetworkCredential(
				userName: Properties.Resources.DocGenerator_AccountName,
				password: Properties.Resources.DocGenerator_Account_Password,
				domain: Properties.Resources.DocGenerator_AccountDomain);
			this.completeDataSet.SDDPdatacontext.MergeOption = MergeOption.NoTracking;

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
			Application.DoEvents();
			}

		private void button1_Click(object sender, EventArgs e)
			{
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
				strExceptionMessage = "*** Exception ERROR ***: DocGeneratorServer cannot access site: " + Properties.Resources.SharePointSiteURL
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
				strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
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
				strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
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
				strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
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
					strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
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
					strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
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
					strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
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
				strExceptionMessage = "*** Exception ERROR ***: DocGeneratorUI cannot access site: " + Properties.Resources.SharePointSiteURL
					+ " Please check that the computer/server is connected to the Domain network "
					+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult + "\nStatusCode: " + exc.StatusCode
					+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
				Application.DoEvents();
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
			catch(DataServiceQueryException)
				{
				Cursor.Current = Cursors.Default;
				MessageBox.Show("The Document Collection ID that you entered doesn't exist. Please enter a valid ID.",
					"Document Collection ID, doesn't exist.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
				}
			catch(DataServiceRequestException exc)
				{
				strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
					+ " Please check that the computer/server is connected to the Domain network "
					+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
					+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
				Application.DoEvents();
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
				strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
					+ " Please check that the computer/server is connected to the Domain network "
					+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
					+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
				Application.DoEvents();
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
					strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
					+ " Please check that the computer/server is connected to the Domain network "
					+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
					+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
					Application.DoEvents();
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
					strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
					+ " Please check that the computer/server is connected to the Domain network "
					+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
					+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
					Application.DoEvents();
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
					strExceptionMessage = "*** Exception ERROR ***: Cannot access site: " + Properties.Resources.SharePointSiteURL
					+ " Please check that the computer/server is connected to the Domain network "
					+ " \n \nMessage:" + exc.Message + "\n HResult: " + exc.HResult
					+ " \nInnerException: " + exc.InnerException + "\nStackTrace: " + exc.StackTrace;
					Application.DoEvents();
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
		}
	}
