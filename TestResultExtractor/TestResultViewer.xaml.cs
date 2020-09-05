#region SystemNamespaces
using System.Data;
using System.Windows;
#endregion

namespace TestResultExtractor
{
    /// <summary>
    /// Interaction logic for TestResultViewer.xaml
    /// </summary>
    public partial class TestResultViewer : Window
    {
        public TestResultViewer()
        {
            InitializeComponent();
        }

        /// <summary>
        /// DisplayData to datagrid
        /// </summary>
        /// <param name="dt"></param>
        public void DisplayData(DataTable dt)
        {
            dataGrid.DataContext = dt.DefaultView;
           
        }
    }
}
