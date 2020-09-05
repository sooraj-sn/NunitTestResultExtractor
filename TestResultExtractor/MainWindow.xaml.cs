
using MahApps.Metro.Controls;

#region SystemNamespaces

using System;
using System.Windows;
#endregion

#region Microsoft Win32 Namespaces
using Microsoft.Win32;
#endregion

namespace TestResultExtractor
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        private string xmlFileName;

        /// <summary>
        /// constructor
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Upload File Event Handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_upload_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = false;
            fileDialog.Filter = "Nunit Xml Files|*.xml";
            fileDialog.DefaultExt = ".xml";
            Nullable<bool> dialogOK = fileDialog.ShowDialog();
            if (dialogOK == true)
            {
                txt_fileupload.Text = fileDialog.FileName;
                txt_fileupload.IsReadOnly = true;
                xmlFileName = txt_fileupload.Text;
                lbl_uploadlabel.Visibility = Visibility.Hidden;
            }
        }
    
        /// <summary>
        /// Export to Excel Event Handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_export_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(xmlFileName))
            {
                lbl_uploadlabel.Visibility = Visibility.Visible;
            }
            else
            {
                lbl_uploadlabel.Visibility = Visibility.Hidden;
                XmlToCSVEngine engine = new XmlToCSVEngine();
                engine.DisplayResultsInExcel(engine.GetTestResultSet(xmlFileName), engine.GetSummaryList(xmlFileName));             
                
                MessageBox.Show("Export Successful");

            }
            

        }

        /// <summary>
        /// View Results in Grid View Event Handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_vwresult_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(xmlFileName))
            {
                lbl_uploadlabel.Visibility = Visibility.Visible;
            }
            else
            {
                lbl_uploadlabel.Visibility = Visibility.Hidden;
                XmlToCSVEngine engine = new XmlToCSVEngine();
                var resultset = engine.GetTestResultSet(xmlFileName);
                TestResultViewer tviewer = new TestResultViewer();
                var dataTable = engine.GetTestResultsInDataTable(resultset);
                tviewer.DisplayData(dataTable);
                tviewer.Show();
            }
           
        }
    }
}
