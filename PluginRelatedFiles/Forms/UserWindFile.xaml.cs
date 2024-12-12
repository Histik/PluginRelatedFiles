using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Shapes;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using System.Globalization;
using System.Reflection;

namespace PluginRelatedFiles.Forms
{
    /// <summary>
    /// Логика взаимодействия для UserWindFile.xaml
    /// </summary>
    public partial class UserWindFile : Window
    {
        private List<string> _listPathFile { get; set; } = new List<string>();

        public UserWindFile(List<string> listPathFile)
        {
            InitializeComponent();
            _listPathFile = listPathFile;
            FileListView.Items.Clear();
            foreach (string path in _listPathFile)
            {
                FileListView.Items.Add(System.IO.Path.GetFileNameWithoutExtension(path));
            }
            if (_listPathFile.Count() == 0)
            {
                FileListView.ItemsSource = new List<string> {"Нету активных связанных файлов!" };
            }
        }
        private void ToFileExportExcel(object sender, RoutedEventArgs e)
        {
            if (_listPathFile.Count() == 0)
            {
                MessageBox.Show("Нету активных связанных файлов!");
                return;
            }
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                string downloadsPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
                string filePath = System.IO.Path.Combine(downloadsPath, "ExportedData.xlsx");
                FileInfo fileInfo = new FileInfo(filePath);

                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet1 = package.Workbook.Worksheets.Add("Лист 1");
                    for (int i = 0; i < _listPathFile.Count(); i++)
                    {
                        worksheet1.Cells[i + 1, 2].Value = _listPathFile[i];
                        worksheet1.Cells[i + 1, 1].Value = System.IO.Path.GetFileNameWithoutExtension(_listPathFile[i]);
                    }

                    package.SaveAs(fileInfo);
                }
                MessageBox.Show("Успешно");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }
            
    }
}
