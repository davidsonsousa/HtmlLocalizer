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
		}

		private void OpenTemplateDialog(object sender, RoutedEventArgs e)
		{
			txtTemplatePath.Text = SystemDialogs.OpenFileDialog();
		}

		private void GenerateLocalizations(object sender, RoutedEventArgs e)
		{
			string[] files = FileManager.GetFilesInFolder(txtExcelFilesPath.Text);

			foreach (string file in files)
			{
				using (var localizer = new Localizer())
				{
					localizer.ProcessTemplate(txtTemplatePath.Text, file);
				}
			}
		}

		private void CloseApplication(object sender, RoutedEventArgs e)
		{
			this.Close();
		}

		void WindowClosing(object sender, CancelEventArgs e)
		{
			var result = MessageBox.Show("Do you really want to exit?", "Html Localizer", MessageBoxButton.YesNo, MessageBoxImage.Question);

			if (result == MessageBoxResult.No)
			{
				e.Cancel = true;
			}
		}
	}
}
