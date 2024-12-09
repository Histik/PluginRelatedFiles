﻿using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using PluginRelatedFiles.Forms;

namespace PluginRelatedFiles
{
    [Transaction(TransactionMode.Manual)]
    public class StartPlugin : IExternalCommand
    {
        private List<string> listPathFile { get; set; } = new List<string>();
        private List<string> listNameFile { get; set; } = new List<string>();
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            List<RevitLinkInstance> linkInstances = collector.OfCategory(BuiltInCategory.OST_RvtLinks)
                                                .WhereElementIsNotElementType()
                                                .ToElements()
                                                .Cast<RevitLinkInstance>()
                                                .ToList();

            foreach (RevitLinkInstance linkInstance in linkInstances)
            {
                listPathFile.Add(linkInstance.GetLinkDocument().PathName);
                listNameFile.Add(System.IO.Path.GetFileName(linkInstance.GetLinkDocument().PathName));
            }

            UserWindFile windFile = new UserWindFile(listPathFile, listNameFile);
            windFile.ShowDialog();

            return Result.Succeeded;
        }
    }
}
