using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.xml;
using System.Windows.Forms;
using System.Collections;
using System.IO;
using System.Data.OleDb;
using System.Threading;
using System.Reflection;


namespace TaxClaimers.WinAppGens
{
	public partial class Form1 : Form
	{

		OpenFileDialog ofd;
		FolderBrowserDialog fbd;
		DataSet ds;
		String selectedServiceProvider;

		enum PHONE_SERVICE_PROVIDER
		{
			Airtel,
			Idea,
			Vodafone,
			ActFiber
		}

		public Form1()
		{
			InitializeComponent();
		}

		private void Form1_Load(object sender, EventArgs e)
		{
			CLearContent();
			updateLabels();
		}

		#region PhoneBills

		private void btnClear_Click(object sender, EventArgs e)
		{
			DateTime d = new DateTime();
			String sdt = "20130925";
			if (DateTime.Compare(d, DateTime.ParseExact(sdt, "yyyyMMdd", null)) > 0)
			{

			}

			CLearContent();
		}

		private void CLearContent()
		{
			ManageLinks(false);
			txtFullName.Clear();
			txtCompany.Clear();
			txtAddressName.Clear();
			txtAddressLine1.Clear();
			txtAddressLine2.Clear();
			txtCity.Clear();
			txtPincode.Clear();
			txtState.Clear();
			txtPhoneNUmber.Clear();
			rdAirtelBills.Checked = false;
			rdIdeaBills.Checked = false;
			rdVodafoneBills.Checked = false;
			UpdatelblDisplay("Welcome...!!");
			txtExcelFile.Clear();
			txtPDFFile.Clear();
			txtLocation.Clear();
			btnProcess.Enabled = false;
		}

		private void btnProcess_Click(object sender, EventArgs e)
		{
			try
			{
				ListFieldNames(txtPDFFile.Text.Trim());
				ImportDataFromExcel(txtExcelFile.Text).First();
				PopulateFiles_UsingExcel();
			}
			catch (Exception ex)
			{
				UpdatelblDisplay(ex.Message);
			}
		}

		private void txtExcelFile_DoubleClick(object sender, EventArgs e)
		{
			ofd = new OpenFileDialog();
			ofd.InitialDirectory = "c:\\";
			ofd.Filter = "Excel files (*.xls)|*.xls|Excel Files (*.xlsx)|*.xlsx";
			if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				txtExcelFile.Text = ofd.FileName;
			}
		}

		private void txtPDFFile_DoubleClick(object sender, EventArgs e)
		{
			ofd = new OpenFileDialog();
			ofd.Filter = "Acrobat files (*.pdf)|*.pdf";
			if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				txtPDFFile.Text = ofd.FileName;
			}
		}

		private void txtLocation_DoubleClick(object sender, EventArgs e)
		{
			fbd = new FolderBrowserDialog();
			if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				txtLocation.Text = fbd.SelectedPath;
			}
		}

		private void rdAirtelBills_CheckedChanged(object sender, EventArgs e)
		{
			GetFileLocations(PHONE_SERVICE_PROVIDER.Airtel);
		}

		private void rdIdeaBills_CheckedChanged(object sender, EventArgs e)
		{
			GetFileLocations(PHONE_SERVICE_PROVIDER.Idea);
		}

		private void rdVodafoneBills_CheckedChanged(object sender, EventArgs e)
		{
			GetFileLocations(PHONE_SERVICE_PROVIDER.Vodafone);
		}

		private void lnkDownloadTemplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
		{
			System.Diagnostics.Process.Start("https://skydrive.live.com/redir?resid=D0473E10EDB339C9!8927&authkey=!AKmDJdb_VNfonpU");
		}

		private void btnRandomnumber_Click(object sender, EventArgs e)
		{
			txtRelationship.Text = GenerateNumber();
		}

		private void PopulateFiles_UsingExcel()
		{
			String newFileName = String.Empty;
			FileInfo fileInfo1 = new FileInfo(Assembly.GetCallingAssembly().Location);
			if (txtLocation.Text == String.Empty)
			{
				txtLocation.Text = fileInfo1.DirectoryName + "\\Resources\\" + txtPhoneNUmber.Text;
			}
			else
			{
				txtLocation.Text = txtLocation.Text + "\\" + txtPhoneNUmber.Text;
			}

			if (System.IO.Directory.Exists(txtLocation.Text))
			{
				System.IO.Directory.Delete(txtLocation.Text, true);
				System.IO.Directory.CreateDirectory(txtLocation.Text);
			}
			else
			{
				System.IO.Directory.CreateDirectory(txtLocation.Text);
			}

			strOpenResultPath = txtLocation.Text;

			for (int i = 1; i < ds.Tables[0].Rows.Count; i++)
			{
				newFileName = "[" + String.Format("{0:00}", i) + "] " + selectedServiceProvider + "_" + txtPhoneNUmber.Text +
								"_" + String.Format("{0:yyyy}", Convert.ToDateTime(ds.Tables[0].Rows[i]["tBillDate"].ToString())) +
								"_" + ds.Tables[0].Rows[i]["Months"].ToString() + ".pdf";

				PdfReader pdfReader = new PdfReader(txtPDFFile.Text.Trim());
				PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(txtLocation.Text + "\\" + newFileName, FileMode.Create));
				AcroFields acroFields = pdfStamper.AcroFields;

				#region "Disabled"
				//acroFields.SetField("FullName", ds.Tables[1].Rows[0]["FullName"].ToString());
				//acroFields.SetField("Company", ds.Tables[1].Rows[0]["Company"].ToString());
				//acroFields.SetField("AddressName", ds.Tables[1].Rows[0]["AddressName"].ToString());
				//acroFields.SetField("AddressLine1", ds.Tables[1].Rows[0]["AddressLine1"].ToString());
				//acroFields.SetField("AddressLine2", ds.Tables[1].Rows[0]["AddressLine2"].ToString());
				//acroFields.SetField("City", ds.Tables[1].Rows[0]["City"].ToString());
				//acroFields.SetField("State", ds.Tables[1].Rows[0]["State"].ToString());
				//acroFields.SetField("Phone", Format2Integer(ds.Tables[1].Rows[0]["Phone"].ToString()));
				//acroFields.SetField("Relationship", Format2Integer(ds.Tables[1].Rows[0]["Relationship"].ToString()));
				#endregion

				acroFields.SetField("tFullName", comboBox1.SelectedItem + " " + txtFullName.Text);
				acroFields.SetField("tCompany", txtCompany.Text);
				acroFields.SetField("tAddressLine1", txtAddressName.Text);
				acroFields.SetField("tAddressLine2", txtAddressLine1.Text);
				acroFields.SetField("tAddressLine3", txtAddressLine2.Text);
				acroFields.SetField("tCityPin", txtCity.Text + "  " + txtPincode.Text);
				acroFields.SetField("tState", txtState.Text);
				acroFields.SetField("tPhoneNo", txtPhoneNUmber.Text);
				acroFields.SetField("tRelationship", txtRelationship.Text);

				acroFields.SetField("tBillNo", Format2Integer(ds.Tables[0].Rows[i]["tBillNo"].ToString()));
				acroFields.SetField("tBillDate", Format2Date(ds.Tables[0].Rows[i]["tBillDate"].ToString()));
				acroFields.SetField("tBillStartPeriod", Format2Date(ds.Tables[0].Rows[i]["tBillStartPeriod"].ToString()));
				acroFields.SetField("tBillEndPeriod", Format2Date(ds.Tables[0].Rows[i]["tBillEndPeriod"].ToString()));
				acroFields.SetField("tDueDate", Format2Date(ds.Tables[0].Rows[i]["tDueDate"].ToString()));
				acroFields.SetField("tPrevBal", Format2Double(ds.Tables[0].Rows[i]["tPrevBal"].ToString()));
				acroFields.SetField("tPrevBal1", Format2Double(ds.Tables[0].Rows[i]["tPrevBal1"].ToString()));
				acroFields.SetField("Adjustments", Format2Double(ds.Tables[0].Rows[i]["Adjustments"].ToString()));
				acroFields.SetField("tFd1", Format2Double(ds.Tables[0].Rows[i]["tFd1"].ToString()));
				acroFields.SetField("tFd2", Format2Double(ds.Tables[0].Rows[i]["tFd2"].ToString()));
				acroFields.SetField("tFd3", Format2Double(ds.Tables[0].Rows[i]["tFd3"].ToString()));
				acroFields.SetField("tFd4", Format2Double(ds.Tables[0].Rows[i]["tFd4"].ToString()));
				acroFields.SetField("tFd5", Format2Double(ds.Tables[0].Rows[i]["tFd5"].ToString()));
				acroFields.SetField("tFd6", Format2Double(ds.Tables[0].Rows[i]["tFd6"].ToString()));
				acroFields.SetField("tFd7", Format2Double(ds.Tables[0].Rows[i]["tFd7"].ToString()));
				acroFields.SetField("tFd8", Format2Double(ds.Tables[0].Rows[i]["tFd8"].ToString()));
				acroFields.SetField("tFd9", Format2Double(ds.Tables[0].Rows[i]["tFd9"].ToString()));
				if (rdIdeaBills.Checked)
					acroFields.SetField("tFd10", Format2Double(ds.Tables[0].Rows[i]["tFd10"].ToString()));
				acroFields.SetField("tFd11", Format2Double(ds.Tables[0].Rows[i]["tFd11"].ToString()));
				acroFields.SetField("tFd12", Format2Double(ds.Tables[0].Rows[i]["tFd12"].ToString()));
				acroFields.SetField("tPages", Format2Integer(ds.Tables[0].Rows[i]["tPages"].ToString()));

				pdfStamper.FormFlattening = true;
				pdfStamper.Close();

				rtbFields.AppendText("Created:" + newFileName);
				rtbFields.AppendText(Environment.NewLine);
				rtbFields.ScrollToCaret();
			}

			UpdatelblDisplay("Job.. Completed. \n Check the files in the following location: " + txtLocation.Text);
			UpdateRTBInfo("Job.. Completed. \n Check the files in the following location: " + txtLocation.Text, true);
			ManageLinks(true);
		}

		private string GenerateFormatedNumber(int jMinValue, int jMaxValue)
		{
			Random r = new Random();
			return r.Next(jMinValue, jMaxValue).ToString();
		}

		private string GenerateFormatedNumber(int j)
		{
			Random random = new Random();
			string r = "";
			int i;
			for (i = 1; i <= j; i++)
			{
				String tr = random.Next(0, 9).ToString();
				if ((i == 1) && (tr == "0"))
					i--;
				else
					r += tr;
			}
			return r;
		}

		private List<DataSet> ImportDataFromExcel(String sInputFileName)
		{
			List<DataSet> _ds = new List<DataSet>();
			string _connString = string.Empty;
			string _Extension = Path.GetExtension(sInputFileName);

			//Checking for the extentions, if XLS connect using Jet OleDB
			if (_Extension.Equals(".xls", StringComparison.CurrentCultureIgnoreCase))
			{
				_connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR=YES;IMEX=0\"";
			}
			//Use ACE OleDb
			else if (_Extension.Equals(".xlsx", StringComparison.CurrentCultureIgnoreCase))
			{
				_connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0;HDR=YES;IMEX=0\"";
			}

			ds = new DataSet();

			using (OleDbConnection oConn = new OleDbConnection(String.Format(_connString, sInputFileName)))
			{
				oConn.Open();
				DataTable dbSchema = oConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables_Info, null);
				foreach (DataRow item in dbSchema.Rows)
				{
					//reading data from excel to Data Table
					using (OleDbCommand oleDbCommand = new OleDbCommand())
					{
						oleDbCommand.Connection = oConn;
						oleDbCommand.CommandText = string.Format("SELECT * FROM [{0}]", item["TABLE_NAME"].ToString());
						using (OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter())
						{
							oleDbDataAdapter.SelectCommand = oleDbCommand;
							oleDbDataAdapter.Fill(ds, item["TABLE_NAME"].ToString());
						}
					}
				}
				_ds.Add(ds);
			}

			return _ds;
		}

		public string GenerateNumber()
		{
			Random random = new Random();
			string r = "";
			int i;
			int numbCount = 0;

			if (rdAirtelBills.Checked)
			{
				numbCount = 11;
			}
			else if (rdIdeaBills.Checked)
			{
				numbCount = 8;
				r = "1.";
			}
			else if (rdVodafoneBills.Checked)
			{
				numbCount = 11;
			}

			for (i = 1; i < numbCount; i++)
			{
				r += random.Next(0, 9).ToString();
			}

			return r;
		}


		private String Format2Date(String sInput)
		{
			return String.Format("{0:dd-MMM-yyyy}", Convert.ToDateTime(sInput));
		}

		private String Format2Integer(String sInput)
		{
			return String.Format("{0:0}", Convert.ToDouble(sInput));
		}

		private String Format2Double(String sInput)
		{
			return String.Format("{0:0.00}", Convert.ToDouble(sInput));
		}

		private void GetFileLocations(PHONE_SERVICE_PROVIDER eServiceProvider)
		{
			FileInfo fileInfo = new FileInfo(Assembly.GetCallingAssembly().Location);
			String sExcelPath, sPDFTemplete;
			Boolean enableCheck = false;

			sExcelPath = fileInfo.DirectoryName + "\\Resources\\" + "Phone_" + eServiceProvider.ToString() + "_Calc.xlsx";
			sPDFTemplete = fileInfo.DirectoryName + "\\Resources\\" + "Phone_" + eServiceProvider.ToString() + "_Template.pdf";

			FileInfo f1 = new FileInfo(sExcelPath);
			FileInfo f2 = new FileInfo(sPDFTemplete);

			if ((!f1.Exists) && !(f2.Exists))
			{
				txtExcelFile.Text = "No file found....";
				txtPDFFile.Text = "No file found....";
				txtLocation.Text = "Unable to set the folder path...";
				enableCheck = false;
			}
			else
			{
				txtRelationship.Text = GenerateNumber();
				txtExcelFile.Text = sExcelPath;
				txtPDFFile.Text = sPDFTemplete;

				selectedServiceProvider = eServiceProvider.ToString();
				enableCheck = true;
			}

			btnProcess.Enabled = enableCheck;
		}

		private void btnPrefill_Click(object sender, EventArgs e)
		{
			comboBox1.SelectedIndex = 0;
			txtFullName.Text = "Raghuram Raichooti";
			txtCompany.Text = "Kanbay Software Ltd";
			txtAddressName.Text = "Plot No 115/32 Nanakram Guda";
			txtAddressLine1.Text = "ISB Road Financial District";
			txtAddressLine2.Text = "Gachibowli";
			txtCity.Text = "Hyderabad";
			txtPincode.Text = "500032";
			txtState.Text = "AP";
			txtPhoneNUmber.Text = "9866183441";
			rdAirtelBills.Checked = false;
			rdIdeaBills.Checked = false;
			rdVodafoneBills.Checked = false;
			UpdatelblDisplay("Welcome...!!");
			txtExcelFile.Clear();
			txtPDFFile.Clear();
			txtLocation.Clear();
		}

		private void txtRelationship_DoubleClick(object sender, EventArgs e)
		{
			bool bStatus;
			bStatus = txtRelationship.ReadOnly;

			if (bStatus)
				txtRelationship.ReadOnly = !bStatus;
		}

		#endregion PhoneBills

		#region PaymentReceipts

		DataSet dsPaymentReceipts;
		string[] AirtelFields = { "AmountPaid", "PaymentDate", "MobileNumber", "PayVia", "AccountNumber", "TransactionRef", "Time", "Name" };
		string[] ActFiberFields = { "DateOfPayment", "PaymentNumber", "DateTime", "AccountNumber", "Amount", "TransRefNumber", "Name", "emailID" };
		string[] IdeaFields = { "PhoneNumber", "AmountText", "ReceiptNo", "Amount", "InvoiceNo", "InvoiceDate", "AccountNo", "ReceiptDate", "Name", "PayVia", "PayVia_Payment" };

		private void txtPRMonthly_Leave(object sender, EventArgs e)
		{
			try
			{
				txtPRYearly.Text = GenerateAmounttoTextBoxes().ToString();
				txtPRAfterDeduction.Text = (Convert.ToDecimal(txtPRYearly.Text) - ((Convert.ToDecimal(txtPRYearly.Text) * 20) / 100)).ToString();
			}
			catch (Exception ex)
			{
				UpdateRTBInfo("validate Monthly textbox" + ex.Message, false);
			}
		}

		private void btnPRRandom_Click(object sender, EventArgs e)
		{
			txtPRAccountNumber.Text = Utilities.GetRandomNumber(Convert.ToInt32(txtPRAccountNumber.Text));
		}

		private void btnPRProcess_Click(object sender, EventArgs e)
		{
			try
			{
				UpdateRTBInfo("Good to start... ", false);
				GetUpdatedAmountFromTextBoxes();
				GenerateData(cmbPRServiceProviders.SelectedItem.ToString());
				GeneratePaymentReceipts(cmbPRServiceProviders.SelectedItem.ToString());
			}
			catch (Exception ex)
			{
				UpdateRTBInfo("found an issue....  ", false);
				UpdateRTBInfo(ex.Message, true);
			}

		}

		private void cmbPRServiceProviders_SelectedIndexChanged(object sender, EventArgs e)
		{
			GetPaymentReceipts(cmbPRServiceProviders.SelectedItem.ToString());
		}

		private void GetPaymentReceipts(string serviceProvider)
		{
			FileInfo fileInfo = new FileInfo(Assembly.GetCallingAssembly().Location);
			String sPDFTemplete = fileInfo.DirectoryName + "\\Resources\\" + "PaymentReceipt_" + serviceProvider + ".pdf";
			Boolean enableCheck = false;

			FileInfo f1 = new FileInfo(sPDFTemplete);

			if (!f1.Exists)
			{
				txtPRTemplatePath.Text = "No file found....";
				UpdateRTBInfo("No template file found...", true);
				enableCheck = false;
			}
			else
			{
				txtPRTemplatePath.Text = sPDFTemplete;
				enableCheck = true;
				ListFieldNames(sPDFTemplete);
			}

			btnPRProcess.Enabled = enableCheck;
		}

		public DataSet GenerateData(string serviceProvider)
		{
			dsPaymentReceipts = new DataSet();
			DataTable dt = new DataTable(serviceProvider);
			Utilities utl = new Utilities();


			switch (serviceProvider)
			{
				case "Airtel":
					foreach (var item in AirtelFields)
					{
						dt.Columns.Add(item);
					}
					dsPaymentReceipts.Tables.Add(dt);

					for (int i = 0; i < 12; i++)
					{
						DataRow dr = dsPaymentReceipts.Tables[serviceProvider].NewRow();

						dr["AmountPaid"] = "Rs " + gennedAmt[i];
						dr["PaymentDate"] = GetPaymentDate(i + 4);
						dr["MobileNumber"] = txtPRMobileNumber.Text.Trim();
						dr["PayVia"] = GetPaymentType();
						dr["AccountNumber"] = txtPRAccountNumber.Text.Trim();
						dr["TransactionRef"] = Utilities.GetRandomNumber(10);
						dr["Time"] = Utilities.GetRandomTime();
						dr["Name"] = txtPRFullName.Text.Trim();

						dsPaymentReceipts.Tables[serviceProvider].Rows.Add(dr);
					}
					break;

				case "Idea":
					foreach (var item in IdeaFields)
					{
						dt.Columns.Add(item);
					}
					dsPaymentReceipts.Tables.Add(dt);

					for (int i = 0; i < 12; i++)
					{
						DataRow dr = dsPaymentReceipts.Tables[serviceProvider].NewRow();
						string amount = gennedAmt[i];
						dr["PhoneNumber"] = txtPRMobileNumber.Text.Trim();
						dr["AmountText"] = "Received a sum of " + utl.NumberToWords(amount) + " .";
						dr["ReceiptNo"] = Utilities.GetRandomNumber(9);
						dr["Amount"] = amount.ToString();
						dr["InvoiceNo"] = Utilities.GetRandomNumber(10);
						dr["InvoiceDate"] = GetInvoiceDate(i + 4);
						dr["AccountNo"] = txtPRAccountNumber.Text.Trim();
						dr["ReceiptDate"] = GetPaymentDate(i + 4);
						dr["Name"] = txtPRFullName.Text.Trim();
						dr["PayVia"] = "VTOPUP";
						dr["PayVia_Payment"] = "VTOPUP_PAYMENT";

						dsPaymentReceipts.Tables[serviceProvider].Rows.Add(dr);
					}
					break;

				case "ActFiber":

					foreach (var item in ActFiberFields)
					{
						dt.Columns.Add(item);
					}
					dsPaymentReceipts.Tables.Add(dt);

					for (int i = 0; i < 12; i++)
					{
						DataRow dr = dsPaymentReceipts.Tables[serviceProvider].NewRow();
						string generatedDate = GetPaymentDate(i + 4);

						dr["DateOfPayment"] = DateTime.Parse(generatedDate).ToString("dd-MMM-yyyy");
						dr["PaymentNumber"] = Utilities.GetRandomNumber(10);
						dr["DateTime"] = DateTime.Parse(generatedDate).ToString("dd MMMM yyyy") + " " + Utilities.GetRandomTime();
						dr["AccountNumber"] = txtPRAccountNumber.Text.Trim();
						dr["Amount"] = gennedAmt[i];
						dr["TransRefNumber"] = Utilities.GetRandomNumber(10);
						dr["Name"] = txtPRFullName.Text.Trim();
						dr["emailID"] = txtPREmailID.Text.Trim();

						dsPaymentReceipts.Tables[serviceProvider].Rows.Add(dr);
					}
					break;

				default:
					break;
			}
			nYMonth = 1;
			niYMonth = 1;
			return dsPaymentReceipts;
		}

		int nYMonth = 1;
		private string GetPaymentDate(int month)
		{
			int rDate;
			int rMonth;
			int rYear = Convert.ToInt32(txtPRFY.Text);

			rDate = Utilities.GetRandomNumberBetween(Convert.ToInt32(cmbPRBilling.SelectedItem), 28);
			rMonth = (month <= 12) ? month : nYMonth++;
			rYear = (nYMonth > 1) ? rYear + 1 : rYear;

			return rDate.ToString() + "/" + rMonth.ToString() + "/" + rYear.ToString();
		}

		int niYMonth = 1;
		private string GetInvoiceDate(int month)
		{
			int rMonth;
			int rYear = Convert.ToInt32(txtPRFY.Text);

			rMonth = (month <= 12) ? month++ : niYMonth++;
			rYear = (niYMonth > 1) ? rYear + 1 : rYear;

			return cmbPRBilling.SelectedItem + "/" + rMonth.ToString() + "/" + rYear.ToString();
		}

		private string GetPaymentType()
		{
			if (rdbPROther.Checked)
				return txtPROther.Text.Trim();

			string[] strPay = cbPRPaidUsing.SelectedItem.ToString().Split(':');

			return strPay[1];
		}

		private void GeneratePaymentReceipts(string ServiceProvider)
		{
			String newFileName = String.Empty;
			List<string> amountList = new List<string>();

			string tSaveLoc = txtPRSaveLocation.Text.Trim();
			string tTemplateLoc = txtPRTemplatePath.Text.Trim();
			string tMobileNumberr = txtPRMobileNumber.Text.Trim();

			FileInfo fileInfo1 = new FileInfo(Assembly.GetCallingAssembly().Location);
			if (tSaveLoc == String.Empty)
			{
				tSaveLoc = fileInfo1.DirectoryName + "\\Resources\\" + tMobileNumberr;
			}
			else
			{
				tSaveLoc = tSaveLoc + "\\" + tMobileNumberr + "_" + cmbPRServiceProviders.SelectedItem.ToString();
			}

			if (System.IO.Directory.Exists(tSaveLoc))
			{
				System.IO.Directory.Delete(tSaveLoc, true);
				System.IO.Directory.CreateDirectory(tSaveLoc);
			}
			else
			{
				System.IO.Directory.CreateDirectory(tSaveLoc);
			}

			strOpenResultPath = tSaveLoc;

			try
			{
				int j = 1;
				for (int i = 0; i < dsPaymentReceipts.Tables[ServiceProvider].Rows.Count; i++)
				{
					newFileName = "[" + String.Format("{0:00}", j) + "] " + liMonths[i].Replace("'", "-") + "-" + ServiceProvider + "_" + tMobileNumberr + ".pdf";

					PdfReader pdfReader = new PdfReader(tTemplateLoc);
					PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(tSaveLoc + "\\" + newFileName, FileMode.Create));
					AcroFields acroFields = pdfStamper.AcroFields;

					foreach (DataColumn dc in dsPaymentReceipts.Tables[ServiceProvider].Columns)
					{
						acroFields.SetField(dc.ColumnName, dsPaymentReceipts.Tables[ServiceProvider].Rows[i][dc.ColumnName].ToString());
					}

					amountList.Add(liMonths[i] + ": " + dsPaymentReceipts.Tables[ServiceProvider].Rows[i][ServiceProvider == "Airtel" ? "AmountPaid" : "Amount"].ToString());

					pdfStamper.FormFlattening = true;
					pdfStamper.Close();
					UpdateRTBInfo("Created:" + newFileName, false);
					j++;
				}
			}
			catch (Exception ex)
			{
				UpdateRTBInfo("found an issue....  ", false);
				UpdateRTBInfo(ex.Message, true);
			}

			UpdateRTBInfo("Generating your bills", true);
			UpdateRTBInfo("Name: " + txtPRFullName.Text, false);
			UpdateRTBInfo("Email ID: " + txtPREmailID.Text, false);
			UpdateRTBInfo("Mobile:" + txtPRMobileNumber.Text, false);
			UpdateRTBInfo("Account No: " + txtPRAccountNumber.Text, false);
			UpdateRTBInfo("_______________________________________", false);
			foreach (var item in amountList)
			{
				UpdateRTBInfo(item, false);
			}

			UpdateRTBInfo("Total Bill Generated for: " + txtPRYearly.Text + " and you will be approved upto " + txtPRAfterDeduction.Text, false);
			UpdateRTBInfo("_______________________________________", false);

			UpdatelblDisplay("Job.. Completed");
			UpdateRTBInfo("Job.. Completed", true);
			ManageLinks(true);
		}

		private void btnPRTemplate_Click(object sender, EventArgs e)
		{
			fbd = new FolderBrowserDialog();
			if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				txtPRTemplatePath.Text = fbd.SelectedPath;
			}
		}

		private void btnPRSave_Click(object sender, EventArgs e)
		{
			fbd = new FolderBrowserDialog();
			if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				txtPRSaveLocation.Text = fbd.SelectedPath;
			}
		}

		Dictionary<string, string> paymentMethods = new Dictionary<string, string>();
		private void LoadPaymentMethods()
		{
			paymentMethods.Add("All", "PayTM");
			paymentMethods.Add("All", "Credit Card");
			paymentMethods.Add("All", "Debit Card");
			paymentMethods.Add("Airtel", "Payment Airtel Money");
			paymentMethods.Add("Idea", "VTOPUP_PAYMENT,VTOPUP");
			paymentMethods.Add("Idea", "BillDeskPayment, WEBPORTAL");
			paymentMethods.Add("ActFiber", "Online Payment");
			paymentMethods.Add("All", "Other");
		}

		List<string> liMonths;

		private void LoadliMonthsList()
		{
			liMonths = new List<string>();

			liMonths.Add("Apr" + txtPRFY.Text.Remove(0, 2));
			liMonths.Add("May" + txtPRFY.Text.Remove(0, 2));
			liMonths.Add("Jun" + txtPRFY.Text.Remove(0, 2));
			liMonths.Add("Jul" + txtPRFY.Text.Remove(0, 2));
			liMonths.Add("Aug" + txtPRFY.Text.Remove(0, 2));
			liMonths.Add("Sep" + txtPRFY.Text.Remove(0, 2));
			liMonths.Add("Oct" + txtPRFY.Text.Remove(0, 2));
			liMonths.Add("Nov" + txtPRFY.Text.Remove(0, 2));
			liMonths.Add("Dec" + txtPRFY.Text.Remove(0, 2));
			liMonths.Add("Jan" + (Convert.ToInt32(txtPRFY.Text.Remove(0, 2)) + 1).ToString());
			liMonths.Add("Feb" + (Convert.ToInt32(txtPRFY.Text.Remove(0, 2)) + 1).ToString());
			liMonths.Add("Mar" + (Convert.ToInt32(txtPRFY.Text.Remove(0, 2)) + 1).ToString());
		}

		Dictionary<int, string> gennedAmt = new Dictionary<int, string>();
		private decimal GenerateAmounttoTextBoxes()
		{
			txtPRApr.Text = ManageGennedAmount(0, false, "0");
			txtPRMay.Text = ManageGennedAmount(1, false, "0");
			txtPRJun.Text = ManageGennedAmount(2, false, "0");
			txtPRJul.Text = ManageGennedAmount(3, false, "0");
			txtPRAug.Text = ManageGennedAmount(4, false, "0");
			txtPRSep.Text = ManageGennedAmount(5, false, "0");
			txtPROct.Text = ManageGennedAmount(6, false, "0");
			txtPRNov.Text = ManageGennedAmount(7, false, "0");
			txtPRDec.Text = ManageGennedAmount(8, false, "0");
			txtPRJan.Text = ManageGennedAmount(9, false, "0");
			txtPRFeb.Text = ManageGennedAmount(10, false, "0");
			txtPRMar.Text = ManageGennedAmount(11, false, "0");

			decimal totalAmount = 0;
			foreach (var item in gennedAmt)
			{
				totalAmount = totalAmount + Convert.ToDecimal(item.Value);
			}

			txtPRYearly.Text = totalAmount.ToString();
			txtPRAfterDeduction.Text = (Convert.ToDecimal(totalAmount) - ((Convert.ToDecimal(totalAmount) * 20) / 100)).ToString();

			return totalAmount;
		}

		private string ManageGennedAmount(int key, bool update, string amount)
		{
			Utilities _utl = new Utilities();

			if (!update)
			{
				amount = Utilities.GetRandomNumberUsingContingency(Convert.ToInt32(txtPRMonthly.Text.Trim()), 300) + "." + Utilities.GetRandomNumberBetween(0, 99);
			}

			UpdateRTBInfo(_utl.NumberToWords(amount), false);

			if (gennedAmt.ContainsKey(key))
				gennedAmt[key] = amount;
			else
				gennedAmt.Add(key, amount);

			return amount.ToString();
		}

		private void GetUpdatedAmountFromTextBoxes()
		{
			try
			{
				ManageGennedAmount(0, true, txtPRApr.Text);
				ManageGennedAmount(1, true, txtPRMay.Text);
				ManageGennedAmount(2, true, txtPRJun.Text);
				ManageGennedAmount(3, true, txtPRJul.Text);
				ManageGennedAmount(4, true, txtPRAug.Text);
				ManageGennedAmount(5, true, txtPRSep.Text);
				ManageGennedAmount(6, true, txtPROct.Text);
				ManageGennedAmount(7, true, txtPRNov.Text);
				ManageGennedAmount(8, true, txtPRDec.Text);
				ManageGennedAmount(9, true, txtPRJan.Text);
				ManageGennedAmount(10, true, txtPRFeb.Text);
				ManageGennedAmount(11, true, txtPRMar.Text);

			}
			catch (Exception)
			{
				UpdateRTBInfo("check textboxes", false);
			}

			decimal totalAmount = 0;
			foreach (var item in gennedAmt)
			{
				totalAmount = totalAmount + Convert.ToDecimal(item.Value);
			}

			txtPRYearly.Text = totalAmount.ToString();
			txtPRAfterDeduction.Text = (Convert.ToInt32(totalAmount) - ((Convert.ToInt32(totalAmount) * 20) / 100)).ToString();
		}

		private void btnPRRefreshAmount_Click(object sender, EventArgs e)
		{
			GenerateAmounttoTextBoxes();
		}

		private void txtPRMonths_Leave(object sender, EventArgs e)
		{
			GetUpdatedAmountFromTextBoxes();
		}

		private void cbPRPaidUsing_SelectedIndexChanged(object sender, EventArgs e)
		{
			if (cbPRPaidUsing.SelectedItem.ToString() == "Other")
			{
				txtPROther.Enabled = true;
			}
			else
			{
				txtPROther.Enabled = false;
			}
		}

		private void txtPRFY_Leave(object sender, EventArgs e)
		{
			updateLabels();
		}

		private void updateLabels()
		{
			int fy = Convert.ToInt32(txtPRFY.Text);

			lblPRFY.Text = "FY-" + txtPRFY.Text + "-" + (Convert.ToInt32(txtPRFY.Text.Remove(0, 2)) + 1).ToString();
			lblPRApr.Text = "Apr'" + txtPRFY.Text.Remove(0, 2);
			lblPRMay.Text = "May'" + txtPRFY.Text.Remove(0, 2);
			lblPRJun.Text = "Jun'" + txtPRFY.Text.Remove(0, 2);
			lblPRJul.Text = "Jul'" + txtPRFY.Text.Remove(0, 2);
			lblPRAug.Text = "Aug'" + txtPRFY.Text.Remove(0, 2);
			lblPRSep.Text = "Sep'" + txtPRFY.Text.Remove(0, 2);
			lblPROct.Text = "Oct'" + txtPRFY.Text.Remove(0, 2);
			lblPRNov.Text = "Nov'" + txtPRFY.Text.Remove(0, 2);
			lblPRDec.Text = "Dec'" + txtPRFY.Text.Remove(0, 2);
			lblPRJan.Text = "Jan'" + (Convert.ToInt32(txtPRFY.Text.Remove(0, 2)) + 1).ToString();
			lblPRFeb.Text = "Feb'" + (Convert.ToInt32(txtPRFY.Text.Remove(0, 2)) + 1).ToString();
			lblPRMar.Text = "Mar'" + (Convert.ToInt32(txtPRFY.Text.Remove(0, 2)) + 1).ToString();

			LoadliMonthsList();
		}
		#endregion

		#region CommonMethods 
		string strOpenResultPath = string.Empty;

		private void ListFieldNames(params string[] filepath)
		{
			PdfReader _pdfReader = new PdfReader(filepath[0]);
			StringBuilder sb = new StringBuilder();
			DictionaryEntry _de = new DictionaryEntry();
			rtbFields.Clear();

			UpdateRTBInfo("List all the fields avaialble in the PDF template.", true);
			foreach (DictionaryEntry de in _pdfReader.AcroFields.Fields)
			{
				_de = de;
				UpdateRTBInfo(_de.Key.ToString(), false);
			}
		}

		private void ManageLinks(Boolean bShow)
		{
			lnkOpenFolder.Visible = bShow;
			lnkDownloadTemplate.Visible = true;
		}

		private void UpdateRTBInfo(String sData, Boolean bLine)
		{
			rtbFields.AppendText(sData);

			if (bLine)
			{
				rtbFields.AppendText(Environment.NewLine);
				rtbFields.AppendText("_______________________________________");
				rtbFields.AppendText(Environment.NewLine);
			}

			rtbFields.AppendText(Environment.NewLine);
			rtbFields.ScrollToCaret();
		}

		private void UpdatelblDisplay(String strInfo)
		{
			lblDisplay.Text = strInfo;
		}

		private void lnkOpenFolder_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
		{
			OpenFileLocation();
		}

		private void OpenFileLocation()
		{
			System.Diagnostics.Process.Start(strOpenResultPath);
		}
		#endregion
	}
}
