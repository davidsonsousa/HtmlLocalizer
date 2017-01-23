using System;
using System.Collections.Generic;
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

			//iterate over the rows and columns and print to the console as it appears in the file
			//excel is not zero based!!
			for (int i = 1; i <= rowCount; i++)
			{
				for (int j = 1; j <= colCount; j++)
				{
					//new line
					if (j == 1)
						Console.Write("\r\n");

					//write the value to the console
					if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
						Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
				}
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
