using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Read_Excel_Table
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

		private void Form1_Load(object sender, EventArgs e)
		{
			// INFO:
			// YOU HAVE TO INSTALL accessdatabaseengine_X64.exe (IT CAN BE DOWNLOAD MICROSOFT WEB SITE, FREE)
			// IF YOU WANT TO USE accessdatabaseengine_X32.exe, YOU SHOULD COMPILE THIS APPLICATION AS 32BIT
		}


		private void button1_Click(object sender, EventArgs e)
		{
			DataTable _dt = GetTable(@GetFileName());

			foreach (DataRow dataRow in _dt.Rows)
			{
				richTextBox1.Text += dataRow.ItemArray[0].ToString() + " - " + dataRow.ItemArray[1].ToString() + "\r\n";
			}
		}


		public static DataTable GetTable(string pFileName)
		{
			if (!File.Exists(pFileName))
			{
				MessageBox.Show("File does not exist!");
				return null;
			}

			var connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0; data source={0}; Extended Properties=Excel 12.0;", pFileName);
			var adapter = new OleDbDataAdapter("SELECT * FROM [Sheet1$]", connectionString);
			var ds = new DataSet();
			string tableName = "Sheet1";

			adapter.Fill(ds, tableName);
			DataTable data = ds.Tables[tableName];
			return data;
		}


		public static string GetFileName()
		{
			Stream myStream;
			OpenFileDialog openFileDialog = new OpenFileDialog();

			openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|Excel 97-2003 files (*.xls)|*.xls|All files (*.*)|*.*";
			openFileDialog.FilterIndex = 1;
			openFileDialog.RestoreDirectory = true;

			try
			{
				if (openFileDialog.ShowDialog() == DialogResult.OK)
				{
					if ((myStream = openFileDialog.OpenFile()) != null)
					{
						myStream.Close();
					}
					return openFileDialog.FileName;
				}
				return "error";
			}
			catch
			{
				MessageBox.Show("Check file, may be it already opened by Excel application!");
				return "error";
			}
		}

	} // CLASS END
}
