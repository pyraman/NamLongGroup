using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace NamLongGroup
{
    public partial class frmDashboard : Form
    {
        private MainMenu mainMenu;

        public frmDashboard()
        {
            mainMenu = new MainMenu();
            MenuItem File = mainMenu.MenuItems.Add("&File");
            File.MenuItems.Add(new MenuItem("&New"));
            File.MenuItems.Add(new MenuItem("&Open"));
            File.MenuItems.Add(new MenuItem("&Export To Excel", new EventHandler(this.ExportToExcel_clicked), Shortcut.CtrlE));
            File.MenuItems.Add(new MenuItem("&Exit"));
            this.Menu = mainMenu;
            MenuItem About = mainMenu.MenuItems.Add("&About");
            About.MenuItems.Add(new MenuItem("&About"));
            this.Menu = mainMenu;
            mainMenu.GetForm().BackColor = Color.Indigo;

            InitializeComponent();
        }

        private void ExportToExcel_clicked(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = @"C:\";
            saveFileDialog.Title = "Save text Files";
            saveFileDialog.CheckFileExists = true;
            saveFileDialog.CheckPathExists = true;
            saveFileDialog.DefaultExt = "xlsx";
            saveFileDialog.Filter = "Text files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            saveFileDialog.FilterIndex = 2;
            saveFileDialog.RestoreDirectory = true;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                String fileName = saveFileDialog.FileName;

                // creating Excel Application  
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                // creating new WorkBook within Excel application  
                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                // creating new Excelsheet in workbook  
                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                // see the excel sheet behind the program  
                app.Visible = true;
                // get the reference of first sheet. By default its name is Sheet1.  
                // store its reference to worksheet  
                worksheet = workbook.Sheets["Sheet1"];
                worksheet = workbook.ActiveSheet;
                // changing the name of active sheet  
                worksheet.Name = "Exported from gridview";
                // storing header part in Excel  
                for (int i = 1; i < gribDashboard.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = gribDashboard.Columns[i - 1].HeaderText;
                }
                // storing Each row and column value to excel sheet  
                for (int i = 0; i < gribDashboard.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < gribDashboard.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = gribDashboard.Rows[i].Cells[j].Value.ToString();
                    }
                }
                // save the application  
                workbook.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                // Exit from the application  
                app.Quit();
            }
        }

        private void frmDashboard_Load(object sender, EventArgs e)
        {

            DataTable table = new DataTable();
            table.Columns.Add(Consts.CUSTOMER);
            table.Columns.Add(Consts.LIABILITY);
            gribDashboard.DataSource = table;

            Double sum = 0;
            DirectoryInfo d = new DirectoryInfo(Consts.DIRECTORY);
            FileInfo[] Files = d.GetFiles("*.xlsx");
            foreach (FileInfo file in Files)
            {
                string sheetName = Consts.WORKSHEET;
                DataTable sheetTable = loadSingleSheet(Consts.DIRECTORY + "\\" + file.Name, sheetName);
                if(sheetTable == null)
                {
                    continue;
                }

                DataRow row = table.NewRow();
                row[Consts.CUSTOMER] = file.Name.Split('.')[0];

                Double value = getData(sheetTable);
                sum += value;
                row[Consts.LIABILITY] = CurrencyFormater.format(value);
                table.Rows.Add(row);
            }


            DataRow rowSum = table.NewRow();
            rowSum[Consts.CUSTOMER] = Consts.SUM;
            rowSum[Consts.LIABILITY] = CurrencyFormater.format(sum);
            table.Rows.Add(rowSum);
        }

        private Double getData(DataTable sheetTable)
        {
            try
            {
                int columnCount = sheetTable.Columns.Count;
                int rowCount = sheetTable.Rows.Count;

                for(int rowIndex = 0; rowIndex < rowCount; rowIndex++)
                {
                    for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
                    {
                        String val = sheetTable.Rows[rowIndex].ItemArray[columnIndex].ToString();
                        if(val.Equals(Consts.LIABILITY) == true)
                        {
                            int foundRow = rowIndex - 1;
                            int foundColumn = columnIndex + 4;

                            String value1 = sheetTable.Rows[foundRow].ItemArray[foundColumn].ToString();
                            String value2 = sheetTable.Rows[foundRow].ItemArray[foundColumn + 1].ToString();

                            return Double.Parse(value1) - Double.Parse(value2);
                        }
                        
                    }
                }

                return 1;
            }
            catch(Exception e)
            {
                return -1;
            }
        }

        private OleDbConnection returnConnection(string filePath)
        {
            return new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + "; Jet OLEDB:Engine Type=5;Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\"");
        }

        private DataTable loadSingleSheet(string filePath, string sheetName)
        {
            if (filePath.Contains("~"))
            {
                return null;
            }
            DataTable sheetData = new DataTable();
            using (OleDbConnection conn = this.returnConnection(filePath))
            {
                conn.Open();
                // retrieve the data using data adapter
                OleDbDataAdapter sheetAdapter = new OleDbDataAdapter("select * from [" + sheetName + "]", conn);
                sheetAdapter.Fill(sheetData);
                conn.Close();
            }

            return sheetData;
        }
    }
}
