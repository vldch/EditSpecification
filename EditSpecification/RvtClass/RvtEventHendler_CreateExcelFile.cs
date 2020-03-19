using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EditSpecification.ExcelClass;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using System.Windows;
using Application = Autodesk.Revit.ApplicationServices.Application;

namespace EditSpecification.RvtClass
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    class RvtEventHendler_CreateExcelFile : IExternalEventHandler
    {

        public Document doc;
        private UIApplication _uiapp;
        public RvtEventHendler_CreateExcelFile(UIApplication uiapp)
        {
            _uiapp = uiapp;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            doc = uidoc.Document;
        }
        public void Execute(UIApplication _uiapp) 
        {
            using (var transact = new Transaction(doc, "Apply new filters to selected Views"))
            {
                transact.Start();
                View view = doc.ActiveView;
                FilteredElementCollector collector = new FilteredElementCollector(view.Document, view.Id);
                List<string> listValue = new List<string>();
                var schedulesheet = view as ViewSchedule;
                var columns = schedulesheet.GetTableData().GetSectionData(SectionType.Body).LastColumnNumber;
                for (int ii = 0; ii < columns + 1; ii++)
                {
                    var header = schedulesheet.GetTableData().GetSectionData(SectionType.Body).GetCellText(0, ii);
                    listValue.Add(header.ToString());
                }
                CreateExcelFile createExcelFile = new CreateExcelFile();
                EditColumnsExcelFile editColumnsExcelFile = new EditColumnsExcelFile();
                createExcelFile.GetExcelFile();
                editColumnsExcelFile.EditExcelFile(listValue);
                TaskDialog.Show("Все ок", "всв  ");
                transact.Commit();
            }

        }
        public string GetName() => nameof(RvtEventHendler_CreateExcelFile);
    }
}
