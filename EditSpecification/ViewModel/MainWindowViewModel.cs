using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using Autodesk.Revit.UI;
using EditSpecification.RvtClass;
using Prism.Commands;
using Prism.Mvvm;

namespace EditSpecification.ViewModel
{
    class MainWindowViewModel
    {
        public ExternalEvent ApplyEvent;
        public MainWindowViewModel()
        {
            CreateExcelFile = new DelegateCommand(createExcelFile);
            EditRevitSpec = new DelegateCommand(editRevitSpec);
        }

        public ICommand CreateExcelFile { get; }
        public ICommand EditRevitSpec { get; }


        private void createExcelFile()
        {
            try
            {
                ApplyEvent.Raise();
            }
            catch
            {
                TaskDialog.Show("Иррор","Це какая то хуйня");
            }
            
        }

        private void editRevitSpec()
        {
            
        }
    }

}
