using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace HtmlLocalizer
{
	public class Localizer : IDisposable
	{
		private Excel.Application _xlApp;
		private Excel.Workbook xlWorkbook;
		private Excel._Worksheet xlWorksheet;
		private Excel.Range xlRange;

		private string _templateName;

		public Localizer()
		{
			_xlApp = new Excel.Application();
		}

		public void ProcessTemplate(string templateFileName, string excelPath)
		{
			try
			{
				// Removes the extension
				_templateName = templateFileName.Remove(templateFileName.LastIndexOf("."));

				var result = OpenSpreadsheet(excelPath);
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		private string LoadTemplate(string path)
		{
			string templateContent = "";

			try
			{
				templateContent = File.ReadAllText(path);
			}
			catch (Exception ex)
			{
				throw ex;
			}

			return templateContent;
		}

		private bool OpenSpreadsheet(string path)
		{
			xlWorkbook = _xlApp.Workbooks.Open(path);

			bool result = LocalizeTemplateFromExcel(path);

			return result;
		}

		private bool LocalizeTemplateFromExcel(string path)
		{
			xlWorksheet = xlWorkbook.Sheets[1];
			xlRange = xlWorksheet.UsedRange;

			int rowCount = xlRange.Rows.Count;
			int colCount = xlRange.Columns.Count;

			Dictionary<string, string> localizedValues = new Dictionary<string, string>();

			// Iterate over the rows and columns. Note: Excel is not zero based
			// Starting with 2 so we ignore the header
			for (int i = 2; i <= rowCount; i++)
			{
				localizedValues.Add(xlRange.Cells[i, 1].Value2.ToString(), xlRange.Cells[i, 3].Value2.ToString());
			}

			return true;
		}

		public void Dispose()
		{
			//cleanup
			GC.Collect();
			GC.WaitForPendingFinalizers();

			//rule of thumb for releasing com objects:
			//  never use two dots, all COM objects must be referenced and released individually
			//  ex: [somthing].[something].[something] is bad

			//release com objects to fully kill excel process from running in the background
			Marshal.ReleaseComObject(xlRange);
			Marshal.ReleaseComObject(xlWorksheet);

			//close and release
			xlWorkbook.Close();
			Marshal.ReleaseComObject(xlWorkbook);

			//quit and release
			_xlApp.Quit();
			Marshal.ReleaseComObject(_xlApp);
		}
	}
}
