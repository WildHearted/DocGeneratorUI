using System;
using System.Collections.Generic;
using System.Net;
using System.Data.Services.Client;
using System.Threading;
using System.Runtime.Remoting.Contexts;
using System.Threading.Tasks;
using DocGeneratorCore;


namespace DocGeneratorUI
	{
	[Synchronization]
	public class DataCache:ContextBoundObject
		{
		CompleteDataSet completeDataSet;

		public void Populate_DataSet(ref CompleteDataSet parDataset)
			{
			
			if(parDataset == null)
				{
				if(this.completeDataSet == null)
					this.completeDataSet = new CompleteDataSet();
				}
			else
				this.completeDataSet = parDataset;

			//- Initiate the SharePoint Datacontext...
			this.completeDataSet.SDDPdatacontext = new DocGeneratorCore.SDDPServiceReference.DesignAndDeliveryPortfolioDataContext(
					new Uri(Properties.Resources.SharePointSiteURL + Properties.Resources.SharePointRESTuri));

			this.completeDataSet.SDDPdatacontext.Credentials = new NetworkCredential(
				userName: Properties.Resources.DocGenerator_AccountName,
				password: Properties.Resources.DocGenerator_Account_Password,
				domain: Properties.Resources.DocGenerator_AccountDomain);
			this.completeDataSet.SDDPdatacontext.MergeOption = MergeOption.NoTracking;

			this.completeDataSet.LastRefreshedOn = new DateTime(2000, 1, 1, 0, 0, 0);
			this.completeDataSet.RefreshingDateTimeStamp = DateTime.UtcNow;
			this.completeDataSet.IsDataSetComplete = false;
			//- --------------------------------------------------------------------------------------------------
			//- Launch the **6 Threads** to build the Complete DataSet while waiting for user input.
			//- --------------------------------------------------------------------------------------------------
			//- These are the 6 threads that will simultaniously populate the Complete DataSet...
			Thread tThread1 = new Thread(() => completeDataSet.PopulateBaseDataObjects());
			Thread tThread2 = new Thread(() => completeDataSet.PopulateBaseDataObjects());
			Thread tThread3 = new Thread(() => completeDataSet.PopulateBaseDataObjects());
			Thread tThread4 = new Thread(() => completeDataSet.PopulateBaseDataObjects());
			Thread tThread5 = new Thread(() => completeDataSet.PopulateBaseDataObjects());
			Thread tThread6 = new Thread(() => completeDataSet.PopulateBaseDataObjects());

			//- Pass the **Thread Number** with each thread start instruction as the parameter to **PopulateDataSet** method.
			tThread1.Name = "Data1";
			tThread1.Start();
			tThread2.Name = "Data2";
			tThread2.Start();
			tThread3.Name = "Data3";
			tThread3.Start();
			tThread4.Name = "Data4";
			tThread4.Start();
			tThread5.Name = "Data5";
			tThread5.Start();
			tThread6.Name = "Data6";
			tThread6.Start();

			//- Lauch the **Synchronisation Thread** which has to wait until all the DataSet Population threads completed,
			//- before it declare the DataSet to be "**Complete**" by setting the **IsDataSetComplete** property.
			Thread threadSychro = new Thread(() => completeDataSet.PopulateBaseDataObjects());
			threadSychro.Name = "Synchro";
			threadSychro.Start();

			//? What happens to the Main thread that was used to call this method?
			//? The main calling thread must wait until the dataset is complete 
			//? Then set the parDataset passed by reference = the Updated Dataset.

			parDataset = this.completeDataSet;
			return;
			}
		}

	}
