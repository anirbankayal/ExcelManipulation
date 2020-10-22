using System.IO;
using System.Windows;
using System.Data;
using Excel;

namespace ExcelManipulation
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void submit_Click(object sender, RoutedEventArgs e)
        {
            string _PathFilename = @"E:\Arghya.xlsx";
            using (FileStream streamIn = File.Open(_PathFilename, FileMode.Open, FileAccess.Read))
            using (IExcelDataReader execlReader = (System.IO.Path.GetExtension(_PathFilename) == ".xlsx" ? ExcelReaderFactory.CreateOpenXmlReader(streamIn) : ExcelReaderFactory.CreateBinaryReader(streamIn)))
            {
                DataSet ds = new DataSet();
                ds = execlReader.AsDataSet();

                int r = ds.Tables[0].Rows.Count;
                int c = ds.Tables[0].Columns.Count;

                MessageBox.Show("Row count is:" +r.ToString());
            }
        }
    }
}
