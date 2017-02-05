using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HtmlLocalizer
{
	public static class PathValidation
	{
		public static ValidationEnum ValidateExcelTextBox(string path)
		{
			if (string.IsNullOrWhiteSpace(path))
			{
				return ValidationEnum.EmptyExcelPath;
			}

			if (Directory.Exists(path) == false)
			{
				return ValidationEnum.InvalidExcelPath;
			}

			return ValidationEnum.Valid;
		}

		public static ValidationEnum ValidateTemplateTextBox(string path)
		{
			if (string.IsNullOrWhiteSpace(path))
			{
				return ValidationEnum.EmptyTemplatePath;
			}

			if (File.Exists(path) == false)
			{
				return ValidationEnum.InvalidTemplatePath;
			}

			return ValidationEnum.Valid;
		}
	}
}
