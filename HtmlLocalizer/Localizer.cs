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

		public Localizer()
		{
			_xlApp = new Excel.Application();
		}

		public string ProcessTemplate(string templateFileName, string excelPath)
		{
			// TODO: Improve this part. Ideally we should let the user select where to save and how to rename the files

			// Replace "excel" by "result" so we can have the result files in a different folder
			string templateName = excelPath.Replace("\\excel\\", "\\result\\");
			// Replace Excel extension (xlsx) by HTML extension (html)
			templateName = templateName.Replace(".xlsx", ".html");

			try
			{
				var excelValues = LoadExcelValues(excelPath);
				var templateText = LoadTemplate(templateFileName);

				foreach (string key in excelValues.Keys)
				{
					templateText = templateText.Replace($"_{key}_", excelValues[key]);
				}

				SaveTemplate(templateName, templateText);
			}
			catch (Exception ex)
			{
				throw ex;
			}

			return $"Localization saved: \"{ templateName }\"";
		}

		private Dictionary<string, string> LoadExcelValues(string path)
		{
			xlWorkbook = _xlApp.Workbooks.Open(path);
			xlWorksheet = xlWorkbook.Sheets[1];
			xlRange = xlWorksheet.UsedRange;

			int rowCount = xlRange.Rows.Count;
			int colCount = xlRange.Columns.Count;

			Dictionary<string, string> excelValues = new Dictionary<string, string>();

			// Iterate over the rows. Note: Excel is not zero based
			// Starting with 2 so we ignore the header
			for (int i = 2; i <= rowCount; i++)
			{
				excelValues.Add(xlRange.Cells[i, 1].Value2.ToString(), xlRange.Cells[i, 3].Value2.ToString());
			}

			return excelValues;
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

		private void SaveTemplate(string path, string text)
		{
			try
			{
				string folder = Path.GetDirectoryName(path);

				if (Directory.Exists(folder) == false)
				{
					Directory.CreateDirectory(folder);
				}

				if (File.Exists(path) == false)
				{
					using (StreamWriter streamWriter = new StreamWriter(path, false, Encoding.UTF8))
					{
						streamWriter.Write(text);
					}
				}
				else
				{
					throw new Exception($"The file {path} already exists.");
				}
			}
			catch (Exception ex)
			{
				throw ex;
			}
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
