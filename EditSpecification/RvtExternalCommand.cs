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
                       
            RvtEventHendler_CreateExcelFile evHendler1 = new RvtEventHendler_CreateExcelFile(doc);
            MainWindowViewModel mainWindowViewModel = new MainWindowViewModel();
            ExternalEvent ExEvent = ExternalEvent.Create(evHendler1);

            
            ExEvent.Raise();
            mainWindowViewModel.ApplyEvent = ExEvent;
            //ExEvent.Initialize();
            //evHendler1.Raise();



            MainWindow mainWindow = new MainWindow(doc);
            mainWindow.Show();
            

            return Result.Succeeded;
        }

    }
}
