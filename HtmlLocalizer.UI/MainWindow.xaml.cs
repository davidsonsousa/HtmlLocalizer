using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace HtmlLocalizer.UI
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();

			btnOpenExcelDialog.Click += OpenExcelDialog;
			btnOpenTemplateDialog.Click += OpenTemplateDialog;
			btnLocalize.Click += GenerateLocalizations;
			btnClose.Click += CloseApplication;

			this.Closing += WindowClosing;
		}

		private void OpenExcelDialog(object sender, RoutedEventArgs e)
		{
			txtExcelFilesPath.Text = SystemDialogs.OpenFolderDialog();
			Log($"Selecting Excel folder: \"{ txtExcelFilesPath.Text }\"");
		}

		private void OpenTemplateDialog(object sender, RoutedEventArgs e)
		{
			txtTemplatePath.Text = SystemDialogs.OpenFileDialog();
			Log($"Selecting Template file: \"{ txtTemplatePath.Text }\"");
		}

		private void GenerateLocalizations(object sender, RoutedEventArgs e)
		{
			Log("Starting localization...");

			try
			{
				string[] files = FileManager.GetFilesInFolder(txtExcelFilesPath.Text);

				foreach (string file in files)
				{
					// Since we are getting all files we should avoid some of them
					if (System.IO.Path.GetFileName(file).StartsWith("~") == false)
					{
						using (var localizer = new Localizer())
						{
							Log(localizer.ProcessTemplate(txtTemplatePath.Text, file));
						}
					}
				}
			}
			catch (Exception ex)
			{
				Log($"Error: {ex.Message}");
			}

			Log("Localization finished");
		}

		private void CloseApplication(object sender, RoutedEventArgs e)
		{
			this.Close();
		}

		private void WindowClosing(object sender, CancelEventArgs e)
		{
			var result = MessageBox.Show("Do you really want to exit?", "Html Localizer", MessageBoxButton.YesNo, MessageBoxImage.Question);

			if (result == MessageBoxResult.No)
			{
				e.Cancel = true;
			}
		}

		private void Log(string message)
		{
			txtLog.AppendText($"[{DateTime.Now}] {message}");
			txtLog.AppendText("\r\n");
		}
	}
}
