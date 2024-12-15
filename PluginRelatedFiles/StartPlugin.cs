using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using PluginRelatedFiles.Forms;
using OfficeOpenXml;
using System.IO;
using System.Windows;
using System;
using Autodesk.Revit.Creation;

namespace PluginRelatedFiles
{
    [Transaction(TransactionMode.Manual)]
    public class StartPlugin : IExternalCommand
    {
        private List<string> listPathFile { get; set; } = new List<string>();
        private List<string> position { get; set; } = new List<string>();
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uiDoc.Document;
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            List<RevitLinkInstance> linkInstances = collector.OfCategory(BuiltInCategory.OST_RvtLinks)
                                                .WhereElementIsNotElementType()
                                                .ToElements()
                                                .Cast<RevitLinkInstance>()
                                                .ToList();

            foreach (RevitLinkInstance linkInstance in linkInstances)
            {
                if (linkInstance.GetLinkDocument() != null)
                {
                    Transform coordinates = linkInstance.GetTransform();
                    listPathFile.Add(linkInstance.GetLinkDocument().PathName);
                    position.Add(coordinates.Origin.ToString());
                }
            }

            UserWindFile windFile = new UserWindFile(listPathFile, position);
            windFile.ShowDialog();

            return Result.Succeeded;
        }
    }
}

namespace PluginRelatedFiles.Forms
{
    /// <summary>
    /// Логика взаимодействия для UserWindFile.xaml
    /// </summary>
    public partial class UserWindFile : Window
    {
        private List<string> _listPathFile { get; set; } = new List<string>();
        private List<string> _position { get; set; } = new List<string>();

        public UserWindFile(List<string> listPathFile, List<string> position)
        {
            InitializeComponent();
            _position = position;
            _listPathFile = listPathFile;
            FileListView.Items.Clear();
            foreach (string path in _listPathFile)
            {
                FileListView.Items.Add(System.IO.Path.GetFileNameWithoutExtension(path));
            }
            if (_listPathFile.Count() == 0)
            {
                FileListView.ItemsSource = new List<string> { "Нету активных связанных файлов!" };
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
                        worksheet1.Cells[i + 1, 3].Value = _position[i];
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


