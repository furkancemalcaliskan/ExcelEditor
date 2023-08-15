using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelEditor.Models
{
	public class DocumentModel
	{
		public string UploadPath { get; set; }
		public string Description { get; set; }
		public string DocumentType { get; set; }
		public string FileSize { get; set; }
	}
}
