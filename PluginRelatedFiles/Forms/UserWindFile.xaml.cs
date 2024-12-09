using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using OfficeOpenXml;

namespace PluginRelatedFiles.Forms
{
    /// <summary>
    /// Логика взаимодействия для UserWindFile.xaml
    /// </summary>
    public partial class UserWindFile : Window
    {
        private List<string> _listPathFile { get; set; } = new List<string>();
        private List<string> _listNameFile { get; set; } = new List<string>();

        public UserWindFile(List<string> listPathFile, List<string> listNameFile)
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _listPathFile = listPathFile;
            _listNameFile = listNameFile;
            FileListView.Items.Clear();
            foreach (string path in _listPathFile)
            {
                FileListView.Items.Add(path);
            }

            foreach (string path in _listNameFile)
            {
                FileListView.Items.Add(path);
            }
        }
        private void ToFileExcel(object sender, EventArgs e)
        {
            CreateExcelFile(_listPathFile, _listNameFile);
        }

        public void CreateExcelFile(List<string> sheet1, List<string> sheet2)
        {
            //string downloadsPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
            //string filePath = System.IO.Path.Combine(downloadsPath, "ExportedData.xlsx");
            string filePath = @"C:\Users\nurma\Downloads\start\ExportedData.xlsx";

            // Создаем новый Excel пакет
            using (var package = new ExcelPackage())
            {
                // Создаем первый лист и заполняем его данными из sheet1
                var worksheet1 = package.Workbook.Worksheets.Add("Sheet1");
                for (int i = 0; i < sheet1.Count(); i++)
                {
                    worksheet1.Cells[i + 1, 1].Value = sheet1[i]; // Заполняем первый столбец
                }

                // Создаем второй лист и заполняем его данными из sheet2
                var worksheet2 = package.Workbook.Worksheets.Add("Sheet2");
                for (int i = 0; i < sheet2.Count(); i++)
                {
                    worksheet2.Cells[i + 1, 1].Value = sheet2[i]; // Заполняем первый столбец
                }

                // Сохраняем файл
                package.SaveAs(new FileInfo(filePath));
            }
        }
    }
}
