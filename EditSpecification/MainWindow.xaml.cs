using System.Windows;
using Autodesk.Revit.DB;

namespace EditSpecification
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public Document document;
        public MainWindow(Document doc)
        {
            document = doc;
            InitializeComponent();
        }
    }
}
