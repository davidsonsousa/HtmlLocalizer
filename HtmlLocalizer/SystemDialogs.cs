using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HtmlLocalizer
{
	public static class SystemDialogs
	{
		public static string OpenFolderDialog()
		{
			var dialog = new FolderBrowserDialog();

			if (dialog.ShowDialog() == DialogResult.OK)
			{
				return dialog.SelectedPath;
			}

			return "";
		}

		public static string OpenFileDialog()
		{
			var dialog = new OpenFileDialog();
			dialog.Filter = "HTML template files (*.html)|*.html|Text files (*.txt)|*.txt";

			if (dialog.ShowDialog() == DialogResult.OK)
			{
				return dialog.FileName;
			}

			return "";
		}
	}
}
