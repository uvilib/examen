using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace WpfApplication2
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
        public static string connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db.mdb;";
        private OleDbConnection myConnection;
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            myConnection = new OleDbConnection(connectString);
            myConnection.Open();

            string query = "SELECT name_author, age_author FROM book";

            OleDbCommand command = new OleDbCommand(query, myConnection);

            command.ExecuteNonQuery();

            OleDbDataAdapter dataAdp = new OleDbDataAdapter(command);
            System.Data.DataTable dt = new System.Data.DataTable("Students");
            dataAdp.Fill(dt);
            dgrid.ItemsSource = dt.DefaultView;
            myConnection.Close();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            excel.Visible = true; //www.yazilimkodlama.com
            Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
            Worksheet sheet1 = (Worksheet)workbook.Sheets[1];

            for (int j = 0; j < dgrid.Columns.Count; j++)
            {
                Range myRange = (Range)sheet1.Cells[1, j + 1];
                sheet1.Cells[1, j + 1].Font.Bold = true;
                sheet1.Columns[j + 1].ColumnWidth = 15;
                myRange.Value2 = dgrid.Columns[j].Header;
            }
            for (int i = 0; i < dgrid.Columns.Count; i++)
            { 
                for (int j = 0; j < dgrid.Items.Count; j++)
                {
                    TextBlock b = dgrid.Columns[i].GetCellContent(dgrid.Items[j]) as TextBlock;
                    Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[j + 2, i + 1];
                    myRange.Value2 = b.Text;
                }
            }

            //Изменить путь
            workbook.SaveAs(@"C:\Users\ianch\Desktop\ПМ03\ggg", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
            false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workbook.Close();
            excel.Quit();
        }
    }
}
