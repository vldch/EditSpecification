using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;
using EditSpecification.RvtClass;
using EditSpecification.ViewModel;
using System.Threading;

namespace EditSpecification
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RvtExternalCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var id = Thread.CurrentThread.ManagedThreadId;
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            View view = doc.ActiveView;

            MainWindowViewModel mainWindowViewModel = new MainWindowViewModel();
            RvtEventHendler_CreateExcelFile evHendler_createExFile = new RvtEventHendler_CreateExcelFile(uiapp);
            RvtEventHendler_EditSpecification evHendler_EditSpecification = new RvtEventHendler_EditSpecification(uiapp);

            ExternalEvent ExEvent_createExFile = ExternalEvent.Create(evHendler_createExFile);
            ExternalEvent ExEvent_editSpec = ExternalEvent.Create(evHendler_EditSpecification);

            mainWindowViewModel.ApplyEvent_createExFile = ExEvent_createExFile;
            mainWindowViewModel.ApplyEvent_editSpec = ExEvent_editSpec;
            MainWindow mainWindow = new MainWindow(doc);
            mainWindow.DataContext = mainWindowViewModel;
            mainWindow.Show();
            return Result.Succeeded;
        }

    }
}
