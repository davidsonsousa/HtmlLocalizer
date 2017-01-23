using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HtmlLocalizer
{
	public class FileManager
	{
		public static string[] GetFilesInFolder(string path)
		{
			string[] files = { };

			if (Directory.Exists(path))
			{
				files = Directory.GetFiles(path, "*.xlsx");
			}

			return files;
		}
	}
}
