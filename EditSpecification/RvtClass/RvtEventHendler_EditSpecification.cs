using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EditSpecification.ExcelClass;

namespace EditSpecification.RvtClass
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    class RvtEventHendler_EditSpecification : IExternalEventHandler
    {
        public Document doc;
        private UIApplication _uiapp;
        public RvtEventHendler_EditSpecification(UIApplication uiapp)
        {
            _uiapp = uiapp;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            doc = uidoc.Document;
        }
        GetExcelChanges getExcelChanges = new GetExcelChanges();
        public void Execute(UIApplication _uiapp)
        {
            var dt = getExcelChanges.GetDataTeble();
            View view = doc.ActiveView;
            using (var transact = new Transaction(doc, "Apply new filters to selected Views"))
            {
                try
                {
                    transact.Start();
                    AddString(view as ViewSchedule);
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        for (int j = 0; j < dt.Rows[i].ItemArray.Length; j++)
                        {
                            string key = dt.Columns[j].ToString();
                            string cellValue = dt.Rows[i].ItemArray[j].ToString();

                            SetCellContent(view as ViewSchedule, GetLastRowNumber(view as ViewSchedule), cellValue, key);
                        }
                        AddString(view as ViewSchedule);
                    }

                    transact.Commit();
                }
                catch
                {
                    TaskDialog.Show("ИРРРОР","Це какая то хуйгя");
                }
                //Нужно считывать параметр элемента а не имя заголовка или чтобы заголовок и параметр были одного имения

            }
        }

        public string GetName() => nameof(RvtEventHendler_CreateExcelFile);
        public void AddString(ViewSchedule schedule)
        {
            TableData calTabledata = schedule.GetTableData();
            TableSectionData tsd = calTabledata.GetSectionData(SectionType.Body);
            tsd.InsertRow(tsd.FirstRowNumber + 1);
        }
        public int GetLastRowNumber(ViewSchedule schedule)
        {
            TableData calTabledata = schedule.GetTableData();
            TableSectionData tsd = calTabledata.GetSectionData(SectionType.Body);
            int lastRowNumber = tsd.LastRowNumber;
            return lastRowNumber;
        }
        public void SetCellContent(View view, int lastRowNimber, string cellValue, string key)
        {
            var schedulesheet = view as ViewSchedule;
            var td = schedulesheet.GetTableData();

            FilteredElementCollector elements = new FilteredElementCollector(view.Document, view.Id);
            List<Element> listOfElement = new List<Element>();
            listOfElement.Clear();

            foreach (Element element in elements)
            {
                listOfElement.Add(element.LookupParameter(key).Element);
            }
            Element setCellCount = listOfElement[listOfElement.Count - 1];
            setCellCount.LookupParameter(key).Set(cellValue);
            double cellValueDouble;
            bool b = double.TryParse(cellValue, out cellValueDouble);
            if (b)
            {
                setCellCount.LookupParameter(key).Set(cellValueDouble);
            }
        }

    }
}
