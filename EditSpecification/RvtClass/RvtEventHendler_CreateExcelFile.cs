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
using Application = Autodesk.Revit.ApplicationServices.Application;

namespace EditSpecification.RvtClass
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    class RvtEventHendler_CreateExcelFile : IExternalEventHandler
    {

        public Document doc;

        public RvtEventHendler_CreateExcelFile(Document Doc)
        {
            doc = Doc;
        }

        //public void Initialize()
        //{
        //    externalEvent = ExternalEvent.Create(this);
        //}
        //public void Raise() => externalEvent.Raise();
        public void Execute(UIApplication uiapp) 
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
                transact.Commit();
            }

        }
        public string GetName() => nameof(RvtEventHendler_CreateExcelFile);
        //public void ExcelOperations()
        //{
        //    CreateExcelFile createExcelFile = new CreateExcelFile();
        //    EditColumnsExcelFile editColumnsExcelFile = new EditColumnsExcelFile();
        //    createExcelFile.GetExcelFile();
        //    editColumnsExcelFile.EditExcelFile(WriteToKeySchedule(view));
        //}
        //public List<string> WriteToKeySchedule(View view)
        //{
        //    FilteredElementCollector collector = new FilteredElementCollector(view.Document, view.Id);
        //    List<string> listValue = new List<string>();
        //    var schedulesheet = view as ViewSchedule;
        //    var columns = schedulesheet.GetTableData().GetSectionData(SectionType.Body).LastColumnNumber;
        //    for (int ii = 0; ii < columns + 1; ii++)
        //    {
        //        var header = schedulesheet.GetTableData().GetSectionData(SectionType.Body).GetCellText(0, ii);
        //        listValue.Add(header.ToString());
        //    }
        //    return listValue;
        //}
    }
}
