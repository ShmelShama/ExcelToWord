using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Windows.Forms;
using System.IO;
using ExcelToWord.Core;
using ExcelToWord.ExcelUtils;
using ExcelToWord.Interfaces;
using ExcelToWord.WordUtils;
using ExcelToWord.Models.ReportModels;
using System.Collections;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToWord.ViewModels
{
    public class MainViewModel: BaseViewModel
    {
        private readonly IReaderFactory _readerFactory;
        private readonly IReportExporter _reportExporter;
        private string _sourceFileName;
        private string _exportPath;
        private string _message;
        private bool _isProcessEnabled;
        public MainViewModel()
        {
            _readerFactory = new ExcelReaderFactory();
            _reportExporter = new ReportWordExporter();
            IsProcessEnabled = false;
        }
        public string SourceFileName
        {
            get=> _sourceFileName;
            set
            {
                _sourceFileName = value;
                OnPropertyChanged(nameof(SourceFileName));
            }
        }
        public string ExportPath
        {
            get => _exportPath;
            set
            {
                _exportPath = value;
                OnPropertyChanged(nameof(ExportPath));
            }
        }
        public string Message
        {
            get => _message;
            set
            {
                _message = value;
                OnPropertyChanged(nameof(Message));
            }
        }

        public bool IsProcessEnabled
        {
            get=>_isProcessEnabled;
            set
            {
                _isProcessEnabled = value;
                OnPropertyChanged(nameof(IsProcessEnabled));
            }
        }
        public RelayCommand BrowseFileCommand => new RelayCommand(o =>
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Excel files|*.xlsx;*.xls;*.xlsb";
            fileDialog.Multiselect = false;
            if(fileDialog.ShowDialog()==DialogResult.Cancel)
                return;
            SourceFileName = fileDialog.FileName;
            CheckEnabled();


        });

        public RelayCommand BrowseFolderCommand => new RelayCommand(o =>
        {
            FolderBrowserDialog folderDialog = new FolderBrowserDialog();
            if (folderDialog.ShowDialog() == DialogResult.Cancel)
                return;
            ExportPath = folderDialog.SelectedPath;
            CheckEnabled();
        });

        public RelayCommand StartProcessCommand => new RelayCommand(async o =>
        {

            Message = "Загрузка данных в программу...";
            var departmentReader = _readerFactory.CreateDepartmentReader();
            var employeeReader = _readerFactory.CreateEmployeeReader();
            var jobReader = _readerFactory.CreateJobReader();

            await Task.Run(() =>
            {
                departmentReader.ReadDataFromFile(SourceFileName);
            });
            if (!departmentReader.Status)
            {
                Message = departmentReader.Message;
                return;
            }
            await Task.Run(() =>
            {
                employeeReader.ReadDataFromFile(SourceFileName);
            });
            if (!employeeReader.Status)
            {
                Message = employeeReader.Message;
                return;
            }
            await Task.Run(() =>
            {
                jobReader.ReadDataFromFile(SourceFileName);
            });
            if (!jobReader.Status)
            {
                Message = jobReader.Message;
                return;
            }
            Message = "Проверка данных...";
            DataIntegrityCheck dataIntegrityCheck = new DataIntegrityCheck();
            await Task.Run(() =>
            {
                dataIntegrityCheck.CheckData(departmentReader.GetData(), employeeReader.GetData(), jobReader.GetData());
            });
            if (!dataIntegrityCheck.Status)
            {
                Message = dataIntegrityCheck.Message;
                return;
            }

            IDepartmentReportCompiler reportCompiler = new DepartmentReportCompiler();
            var reportData = reportCompiler.CreateReport(departmentReader.GetData(), employeeReader.GetData(), jobReader.GetData());
            if (reportData == null)
            {
                Message = "Не удалось подготовить данные для отчета";
                return;
            }
            await Task.Run(() =>
            {
                _reportExporter.SetExportData(reportData);
                _reportExporter.Export(ExportPath);
            });

            Message = _reportExporter.Message;


        });

        public void CheckEnabled()
        {
            if(string.IsNullOrWhiteSpace(SourceFileName) || string.IsNullOrWhiteSpace(ExportPath))
            {
                IsProcessEnabled = false;
                return;
            }
            IsProcessEnabled= true;
             
            
        }
    }
}
