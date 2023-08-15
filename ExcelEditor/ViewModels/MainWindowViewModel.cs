using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Windows.Input;
using ExcelEditor.Models;
using Microsoft.Win32;
using System.Runtime.CompilerServices;
using System.Collections.Generic;
using OfficeOpenXml;
using System.Runtime.InteropServices;
using System.IO.Compression;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;
using System.Windows.Threading;
using System.Linq;

namespace ExcelEditor.ViewModels
{
    public class MainWindowViewModel : INotifyPropertyChanged
    {
        private double _progressValue;
        private bool _isProgressBarVisible;

        public double ProgressValue
        {
            get => _progressValue;
            set
            {
                _progressValue = value;
                OnPropertyChanged();
            }
        }

        public bool IsProgressBarVisible
        {
            get => _isProgressBarVisible;
            set
            {
                _isProgressBarVisible = value;
                OnPropertyChanged();
            }
        }

        private bool _isConvertButtonEnabled;
        private bool _isAddButtonEnabled = true;
        private bool _isEditButtonEnabled;
        private string _messageBoxText;
        private ObservableCollection<DocumentModel> _documentList = new ObservableCollection<DocumentModel>();

        public bool IsConvertButtonEnabled
        {
            get => _isConvertButtonEnabled;
            set
            {
                _isConvertButtonEnabled = value;
                OnPropertyChanged();
            }
        }

        public bool IsAddButtonEnabled
        {
            get => _isAddButtonEnabled;
            set
            {
                _isAddButtonEnabled = value;
                OnPropertyChanged();
            }
        }

        public bool IsEditButtonEnabled
        {
            get => _isEditButtonEnabled;
            set
            {
                _isEditButtonEnabled = value;
                OnPropertyChanged();
            }
        }

        public ObservableCollection<DocumentModel> DocumentList
        {
            get => _documentList;
            set
            {
                _documentList = value;
                OnPropertyChanged();
            }
        }

        public string MessageBoxText
        {
            get => _messageBoxText;
            set
            {
                _messageBoxText = value;
                OnPropertyChanged();
            }
        }

        private ICommand _onAddDocumentCommand;
        public ICommand OnAddDocumentCommand
        {
            get
            {
                if (_onAddDocumentCommand == null)
                    _onAddDocumentCommand = new DelegateCommand(AddDocument);
                return _onAddDocumentCommand;
            }
        }

        private ICommand _onEditDocumentCommand;
        public ICommand OnEditDocumentCommand
        {
            get
            {
                if (_onEditDocumentCommand == null)
                    _onEditDocumentCommand = new DelegateCommand(EditDocument);
                return _onEditDocumentCommand;
            }
        }

        private ICommand _convertCommand;
        public ICommand ConvertCommand
        {
            get
            {
                if (_convertCommand == null)
                    _convertCommand = new DelegateCommand(Convert);
                return _convertCommand;
            }
        }

        private void ShowMessageBox(string message)
        {
            MessageBoxText = message;
        }

        public void AddDocument(object parameter)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Zip Files|*.zip";
            openFileDialog.Multiselect = false;

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                FileInfo fileInfo = new FileInfo(filePath);
                double fileSize = (double)fileInfo.Length / (1024 * 1024); // Convert to MB

                DocumentModel newDocument = new DocumentModel
                {
                    UploadPath = filePath,
                    Description = Path.GetFileName(filePath),
                    DocumentType = Path.GetExtension(filePath),
                    FileSize = fileSize.ToString("F2") + " MB"
                };

                DocumentList.Add(newDocument);

                IsConvertButtonEnabled = true;
                IsAddButtonEnabled = false;
                IsEditButtonEnabled = true;

                ShowMessageBox("Döküman Başarıyla Eklendi!");
            }
        }

        private void EditDocument(object parameter)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Zip Files|*.zip";
            openFileDialog.Multiselect = false;

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                FileInfo fileInfo = new FileInfo(filePath);
                double fileSize = (double)fileInfo.Length / (1024 * 1024); // Convert to MB

                DocumentModel editedDocument = new DocumentModel
                {
                    UploadPath = filePath,
                    Description = Path.GetFileName(filePath),
                    DocumentType = Path.GetExtension(filePath),
                    FileSize = fileSize.ToString("F2") + " MB"
                };

                if (DocumentList.Count > 0)
                {
                    DocumentList[DocumentList.Count - 1] = editedDocument;
                }

                IsConvertButtonEnabled = true;
                IsAddButtonEnabled = false;
                IsEditButtonEnabled = true;

                ShowMessageBox("Döküman Başarıyla Güncellendi!");
            }
        }

        private async void Convert(object parameter)
        {
            IsConvertButtonEnabled = false;
            IsAddButtonEnabled = false;
            IsEditButtonEnabled = false;

            ShowMessageBox("Dönüştürme İşlemi Başlatıldı.");

            if (DocumentList.Count == 0)
            {
                IsProgressBarVisible = false;
                ShowMessageBox("Dönüştürülecek döküman bulunamadı.");
                return;
            }

            IsProgressBarVisible = true;
            ProgressValue = 0;

            int totalCount = DocumentList.Count;
            int processedCount = 0;

            var progress = new Progress<double>(percentage => ProgressValue = percentage);

            await Task.Run(() =>
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook combinedWorkbook = excelApp.Workbooks.Add();
                Excel.Worksheet combinedWorksheet = combinedWorkbook.Worksheets.Add();

                foreach (var document in DocumentList)
                {
                    if (Path.GetExtension(document.UploadPath).ToLower() == ".zip")
                    {
                        string tempFolder = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
                        Directory.CreateDirectory(tempFolder);

                        ZipFile.ExtractToDirectory(document.UploadPath, tempFolder);

                        List<string> excelFilePaths = Directory.GetFiles(tempFolder, "*.xls").ToList();

                        int totalExcelFiles = excelFilePaths.Count;
                        int processedExcelFiles = 0;

                        foreach (var filePath in excelFilePaths)
                        {
                            double progressPercentage = (double)(processedCount * totalExcelFiles + processedExcelFiles) / (totalCount * totalExcelFiles) * 100;
                            ((IProgress<double>)progress).Report(progressPercentage);

                            Excel.Workbook sourceWorkbook = excelApp.Workbooks.Open(filePath);

                            foreach (Excel.Worksheet sourceWorksheet in sourceWorkbook.Worksheets)
                            {
                                int lastRow = combinedWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 1;
                                CopyData(sourceWorksheet, combinedWorksheet, lastRow);
                            }

                            sourceWorkbook.Close(false);
                            Marshal.ReleaseComObject(sourceWorkbook);

                            processedExcelFiles++;
                        }

                        processedCount++;
                    }
                }

                string combinedExcelPath = Path.Combine(Path.GetDirectoryName(DocumentList[0].UploadPath), "CombinedExcel.xls");
                combinedWorkbook.SaveAs(combinedExcelPath, Excel.XlFileFormat.xlWorkbookNormal);

                combinedWorkbook.Close(false);
                Marshal.ReleaseComObject(combinedWorkbook);
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);

                // Clean temp folders
                CleanTempFoldersAsync(DocumentList).Wait();
            });

            IsProgressBarVisible = false;
            ProgressValue = 0;

            IsConvertButtonEnabled = true;
            IsAddButtonEnabled = false;
            IsEditButtonEnabled = true;

            ShowMessageBox("Dönüşüm başarıyla tamamlandı ve dosyalar birleştirildi.");
        }





        private void CopyData(Excel.Worksheet sourceWorksheet, Excel.Worksheet destinationWorksheet, int startRow)
        {
            int lastRow = sourceWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            int lastCol = sourceWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

            for (int row = 1; row <= lastRow; row++)
            {
                for (int col = 1; col <= lastCol; col++)
                {
                    destinationWorksheet.Cells[startRow + row - 1, col] = sourceWorksheet.Cells[row, col];
                }
            }
        }

        private async Task CleanTempFoldersAsync(ObservableCollection<DocumentModel> documents)
        {
            foreach (var document in documents)
            {
                if (Path.GetExtension(document.UploadPath).ToLower() == ".zip")
                {
                    string tempFolder = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(document.UploadPath));
                    if (Directory.Exists(tempFolder))
                    {
                        Directory.Delete(tempFolder, true);
                        await Task.Delay(100); // Wait a bit to ensure the folder is deleted
                    }
                }
            }
        }
        // Other methods and properties here

        public event PropertyChangedEventHandler PropertyChanged;

		protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
		{
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
		}
	}

	public class DelegateCommand : ICommand
	{
		private readonly Action<object> _execute;
		private readonly Func<object, bool> _canExecute;

		public event EventHandler CanExecuteChanged;

		public DelegateCommand(Action<object> execute, Func<object, bool> canExecute = null)
		{
			_execute = execute ?? throw new ArgumentNullException(nameof(execute));
			_canExecute = canExecute;
		}

		public bool CanExecute(object parameter)
		{
			return _canExecute == null || _canExecute(parameter);
		}

		public void Execute(object parameter)
		{
			_execute(parameter);
		}

		public void RaiseCanExecuteChanged()
		{
			CanExecuteChanged?.Invoke(this, EventArgs.Empty);
		}
	}
}
